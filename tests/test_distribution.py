import importlib.util
import io
import json
from pathlib import Path
import subprocess
import tempfile
import unittest
import zipfile

ROOT = Path(__file__).resolve().parents[1]


def load(name):
    spec = importlib.util.spec_from_file_location(name, ROOT / 'tools' / (name + '.py'))
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


d = load('distribute_skills')
doctor = load('skill_doctor')


class DistributionTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.base = Path(self.tmp.name)
        self.repo = self.base / 'repo'
        self.repo.mkdir()
        subprocess.run(['git', 'init', '-q', str(self.repo)], check=True)
        self.old_tool_root = d.TOOL_ROOT
        d.TOOL_ROOT = self.repo
        (self.repo / 'tools').mkdir()
        (self.repo / 'tools/skill_doctor.py').write_bytes((ROOT / 'tools/skill_doctor.py').read_bytes())
        (self.repo / 'SKILL.md').write_text('---\nname: sample-skill\ndescription: Test\n---\nRead rules.md\n')
        (self.repo / 'rules.md').write_text('Stop when evidence is missing.\n')
        (self.repo / 'excluded.txt').write_text('not distributed')
        manifest = {'name': 'sample-skill', 'version': '1.0.0', 'root': '.',
                    'repository': 'https://example.invalid/sample', 'visibility': 'public',
                    'include': ['SKILL.md', 'rules.md'], 'prompt': ['SKILL.md', 'rules.md'],
                    'stages': {'rules': ['file:SKILL.md']}, 'limitations': ['No execution proof']}
        (self.repo / 'skill-package.json').write_text(json.dumps(manifest))
        self.commit()

    def commit(self):
        subprocess.run(['git', '-C', str(self.repo), 'add', '.'], check=True)
        subprocess.run(['git', '-C', str(self.repo), '-c', 'user.name=Test', '-c',
                        'user.email=test@example.invalid', 'commit', '-qm', 'fixture'], check=True)

    def tearDown(self):
        d.TOOL_ROOT = self.old_tool_root
        self.tmp.cleanup()

    def build(self, folder='out'):
        out = self.base / folder
        result = d.build(self.repo, 'HEAD', out)
        return out / 'sample-skill-1.0.0.skill.zip', result

    def test_deterministic_committed_allowlist_and_install(self):
        (self.repo / 'rules.md').write_text('uncommitted edit must not ship')
        first, result = self.build()
        second, _ = self.build('again')
        self.assertEqual(first.read_bytes(), second.read_bytes())
        name, files = d.unpack_checked(first.read_bytes())
        self.assertNotIn('excluded.txt', files)
        rules = (first.parent / 'sample-skill-1.0.0.rules.txt').read_bytes()
        self.assertIn(b'Stop when evidence', rules)
        self.assertNotIn(b'uncommitted edit', rules)
        d.install(first, result['zip_sha256'], self.base / 'installed')
        self.assertEqual((self.base / 'installed' / name / 'rules.md').read_bytes(), files['rules.md'])
        with self.assertRaises(ValueError):
            d.install(first, result['zip_sha256'], self.base / 'installed')

    def test_wrong_external_checksum(self):
        archive, _ = self.build()
        with self.assertRaises(ValueError):
            d.install(archive, '0' * 64, self.base / 'install')
        self.assertFalse((self.base / 'install').exists())

    def test_tampered_member(self):
        archive, _ = self.build()
        _, files = d.unpack_checked(archive.read_bytes())
        files['rules.md'] = b'changed'
        data = io.BytesIO()
        with zipfile.ZipFile(data, 'w') as z:
            for path, content in files.items():
                z.writestr('sample-skill/' + path, content)
        with self.assertRaises(ValueError):
            d.unpack_checked(data.getvalue())

    def test_zip_slip_and_symlink(self):
        for name, mode in [('sample-skill/../escape', 0o100644), ('sample-skill/link', 0o120777)]:
            data = io.BytesIO()
            with zipfile.ZipFile(data, 'w') as z:
                info = zipfile.ZipInfo(name)
                info.external_attr = mode << 16
                z.writestr(info, b'bad')
            with self.assertRaises(ValueError):
                d.unpack_checked(data.getvalue())

    def test_missing_include_fails(self):
        path = self.repo / 'skill-package.json'
        manifest = json.loads(path.read_text())
        manifest['include'].append('missing/*')
        path.write_text(json.dumps(manifest))
        self.commit()
        with self.assertRaises(ValueError):
            self.build()

    def test_doctor_stage_isolation(self):
        manifest = {'name': 'sample-skill', 'stages': {'rules': ['file:SKILL.md'],
                    'render': ['bin:nonexistent-skill-test-executable']}, 'limitations': []}
        result = doctor.check(self.repo, manifest, 'rules')
        self.assertTrue(result['stages']['rules']['dependencies_found'])
        self.assertEqual(result['functional_test'], 'not_run')
        self.assertFalse(doctor.check(self.repo, manifest)['stages']['render']['dependencies_found'])
        with self.assertRaises(ValueError):
            doctor.check(self.repo, manifest, 'made-up')


if __name__ == '__main__':
    unittest.main()
