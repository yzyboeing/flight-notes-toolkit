#!/usr/bin/env python3
"""Read-only dependency probes. No installs, credential reads, or QA claims."""
import argparse
import importlib.util
import json
from pathlib import Path
import shutil
import subprocess
import sys


def probe(item, root):
    kind, name = item.split(':', 1)
    if kind == 'bin':
        return shutil.which(name) is not None
    if kind == 'py':
        try:
            return importlib.util.find_spec(name) is not None
        except (ImportError, ValueError):
            return False
    if kind == 'node':
        if not shutil.which('node'):
            return False
        try:
            return subprocess.run(['node', '-e', 'require(process.argv[1])', name],
                                  cwd=root, stdout=subprocess.DEVNULL,
                                  stderr=subprocess.DEVNULL, timeout=15).returncode == 0
        except (OSError, subprocess.TimeoutExpired):
            return False
    if kind == 'file':
        return (root / name).is_file()
    raise ValueError('Unknown probe type')


def check(root, manifest, stage=None):
    stages = manifest['stages']
    if stage and stage not in stages:
        raise ValueError('Unknown stage; choose: ' + ', '.join(stages))
    report = {}
    for key, requirements in stages.items():
        if stage and stage != key:
            continue
        probes = {'python>=3.10': sys.version_info >= (3, 10)}
        probes.update({item: probe(item, root) for item in requirements})
        report[key] = {'dependencies_found': all(probes.values()), 'probes': probes}
    return {'skill': manifest['name'], 'python': '.'.join(map(str, sys.version_info[:3])),
            'stages': report, 'functional_test': 'not_run', 'visual_qa': 'not_run',
            'limitations': manifest['limitations'],
            'next_checks': manifest.get('next_checks', [])}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--root', type=Path, default=Path(__file__).resolve().parent)
    parser.add_argument('--stage')
    args = parser.parse_args()
    try:
        root = args.root.resolve()
        manifest = json.loads((root / 'skill-package.json').read_text())
        result = check(root, manifest, args.stage)
    except (OSError, ValueError, KeyError):
        print(json.dumps({'error': 'Invalid manifest, package root, or stage'}))
        return 2
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if all(s['dependencies_found'] for s in result['stages'].values()) else 1


if __name__ == '__main__':
    sys.exit(main())
