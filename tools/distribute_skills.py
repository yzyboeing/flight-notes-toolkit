#!/usr/bin/env python3
"""Build immutable-source skill bundles and install verified ZIPs without overwrites."""
import argparse
import fnmatch
import hashlib
import io
import json
from pathlib import Path, PurePosixPath
import re
import shutil
import stat
import subprocess
import tempfile
import zipfile


TOOL_ROOT = Path(__file__).resolve().parents[1]


def git(repo, *args):
    return subprocess.check_output(['git', '-C', str(repo), *args])


def digest(data):
    return hashlib.sha256(data).hexdigest()


def valid_path(name):
    p = PurePosixPath(name)
    return (bool(name) and not p.is_absolute() and '\\' not in name and ':' not in name
            and all(x not in ('', '.', '..') for x in name.split('/')))


def selected(name, patterns):
    return any(fnmatch.fnmatchcase(name, p) for p in patterns)


def build(repo, ref, out):
    if git(TOOL_ROOT, 'diff', 'HEAD', '--', 'tools/distribute_skills.py', 'tools/skill_doctor.py'):
        raise ValueError('Commit builder changes before building a release')
    commit = git(repo, 'rev-parse', '--verify', ref + '^{commit}').decode().strip()
    manifest = json.loads(git(repo, 'show', commit + ':skill-package.json'))
    name, version = manifest['name'], manifest['version']
    if not re.fullmatch(r'[a-z0-9]+(?:-[a-z0-9]+)*', name):
        raise ValueError('Invalid skill name')
    if not re.fullmatch(r'\d+\.\d+\.\d+', version):
        raise ValueError('Invalid package version')
    if manifest['root'] != '.' and not valid_path(manifest['root']):
        raise ValueError('Invalid source root')
    prefix = '' if manifest['root'] == '.' else manifest['root']
    prefix = prefix + '/' if prefix else ''
    payload = {}
    for entry in git(repo, 'ls-tree', '-r', '-z', commit).split(b'\0'):
        if not entry:
            continue
        meta, path = entry.split(b'\t', 1)
        path = path.decode()
        if not path.startswith(prefix):
            continue
        rel = path[len(prefix):]
        if not selected(rel, manifest['include']):
            continue
        if not valid_path(rel) or meta.split()[0] not in (b'100644', b'100755'):
            raise ValueError('Unsafe path or non-regular file in bundle')
        payload[rel] = git(repo, 'show', commit + ':' + path)
    if 'SKILL.md' not in payload:
        raise ValueError('SKILL.md missing')
    match = re.search(rb'^name:\s*[\x22\x27]?([^\r\n\x22\x27]+)', payload['SKILL.md'], re.M)
    if not match or match[1].decode().strip() != name:
        raise ValueError('Skill name does not match manifest')
    for pattern in manifest['include']:
        if not any(fnmatch.fnmatchcase(p, pattern) for p in payload):
            raise ValueError('Empty include pattern: ' + pattern)
    for target, source_path in manifest.get('extra_files', {}).items():
        if not valid_path(target) or not valid_path(source_path) or target in payload:
            raise ValueError('Invalid extra file mapping')
        payload[target] = git(repo, 'show', commit + ':' + source_path)
    # Builder and doctor always come from one committed toolkit revision.
    tool_commit = git(TOOL_ROOT, 'rev-parse', 'HEAD').decode().strip()
    doctor = git(TOOL_ROOT, 'show', tool_commit + ':tools/skill_doctor.py')
    payload['doctor.py'] = doctor
    payload['skill-package.json'] = (json.dumps(manifest, ensure_ascii=False, indent=2) + '\n').encode()
    header = (f'# {name} rules bundle {version}\n\n'
              f'Source: {manifest["repository"]} @ {commit}\n\n'
              'Generated, not a second source. Read SKILL.md first and follow its routing; '
              'later sections are named reference files, not instructions to execute every workflow. '
              'Relative links refer to those file names. This text cannot grant tools, repository '
              'access, publication authority, or approval of unseen artifacts. If the platform '
              'cannot read this complete bundle, report missing sections. Scripts and assets '
              'are in the matching ZIP. Private rules must remain in authorized environments.\n\n')
    prompt_files = [p for p in payload if selected(p, manifest['prompt'])]
    prompt_files.sort(key=lambda p: (p != 'SKILL.md', p))
    for pattern in manifest['prompt']:
        if not any(fnmatch.fnmatchcase(p, pattern) for p in prompt_files):
            raise ValueError('Empty prompt pattern: ' + pattern)
    prompt = header + '\n'.join(f'\n<!-- FILE: {p} -->\n{payload[p].decode()}\n<!-- END FILE: {p} -->\n'
                               for p in prompt_files)
    source = {'schema_version': 1, 'repository': manifest['repository'], 'commit': commit,
              'name': name, 'version': version, 'visibility': manifest['visibility'],
              'builder_repository': 'https://github.com/yzyboeing/flight-notes-toolkit',
              'builder_commit': tool_commit,
              'files': {p: digest(data) for p, data in sorted(payload.items())}}
    payload['SOURCE.json'] = (json.dumps(source, indent=2) + '\n').encode()
    out.mkdir(parents=True, exist_ok=True)
    stem = f'{name}-{version}'
    zip_path, text_path = out / (stem + '.skill.zip'), out / (stem + '.rules.txt')
    checksum_path = out / (stem + '.sha256')
    if any(p.exists() for p in (zip_path, text_path, checksum_path)):
        raise ValueError('Output exists; use a new output directory')
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as archive:
        for path, data in sorted(payload.items()):
            info = zipfile.ZipInfo(name + '/' + path, (2020, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = (stat.S_IFREG | 0o644) << 16
            archive.writestr(info, data)
    text_path.write_bytes(prompt.encode())
    checksum_path.write_text(''.join(f'{digest(p.read_bytes())}  {p.name}\n' for p in (zip_path, text_path)))
    return {'skill': name, 'commit': commit, 'version': version, 'files': len(payload),
            'zip_sha256': digest(zip_path.read_bytes())}


def unpack_checked(data):
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        infos = archive.infolist()
        if len(infos) > 2000 or sum(i.file_size for i in infos) > 100_000_000:
            raise ValueError('Archive exceeds skill bundle limits')
        names = [i.filename for i in infos]
        if len(names) != len(set(names)):
            raise ValueError('Duplicate archive entries')
        if any(not valid_path(n) for n in names):
            raise ValueError('Unsafe archive path')
        roots = {n.split('/')[0] for n in names}
        if len(roots) != 1:
            raise ValueError('Bundle must contain one skill directory')
        root = roots.pop()
        if not re.fullmatch(r'[a-z0-9]+(?:-[a-z0-9]+)*', root):
            raise ValueError('Invalid root name')
        if any(stat.S_IFMT(i.external_attr >> 16) not in (0, stat.S_IFREG) for i in infos):
            raise ValueError('Links and special files are not accepted')
        files = {n[len(root) + 1:]: archive.read(n) for n in names}
    source = json.loads(files['SOURCE.json'])
    if source['name'] != root or set(source['files']) != set(files) - {'SOURCE.json'}:
        raise ValueError('Source manifest file set mismatch')
    if any(digest(files[p]) != sha for p, sha in source['files'].items()):
        raise ValueError('Source manifest checksum mismatch')
    if 'SKILL.md' not in files or 'doctor.py' not in files:
        raise ValueError('Incomplete skill bundle')
    return root, files


def install(archive_path, expected_sha256, destination):
    data = archive_path.read_bytes()
    if digest(data) != expected_sha256.lower():
        raise ValueError('Archive checksum mismatch')
    name, files = unpack_checked(data)
    destination = destination.expanduser().resolve()
    target = destination / name
    if target.exists() or target.is_symlink():
        raise ValueError('Destination exists; back up the exact old installation outside discovery first')
    destination.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='.skill-stage-', dir=destination) as tmp:
        stage = Path(tmp) / name
        stage.mkdir()
        for path, contents in files.items():
            file = stage / path
            file.parent.mkdir(parents=True, exist_ok=True)
            file.write_bytes(contents)
        stage.rename(target)
    return {'installed': name, 'sha256_verified': True, 'dependencies_installed': False}


def main():
    p = argparse.ArgumentParser(description=__doc__)
    subs = p.add_subparsers(dest='command', required=True)
    b = subs.add_parser('build')
    b.add_argument('--repo', type=Path, required=True)
    b.add_argument('--ref', default='HEAD')
    b.add_argument('--out', type=Path, required=True)
    i = subs.add_parser('install')
    i.add_argument('archive', type=Path)
    i.add_argument('--sha256', required=True)
    i.add_argument('--destination', type=Path, required=True)
    args = p.parse_args()
    try:
        result = (build(args.repo, args.ref, args.out) if args.command == 'build'
                  else install(args.archive, args.sha256, args.destination))
    except (OSError, ValueError, KeyError, subprocess.CalledProcessError, zipfile.BadZipFile) as e:
        p.exit(1, 'Distribution failed: ' + str(e) + '\n')
    print(json.dumps(result, indent=2))


if __name__ == '__main__':
    main()
