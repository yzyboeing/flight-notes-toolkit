from pathlib import Path
import json
import subprocess
import sys
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]


class PackageTest(unittest.TestCase):
    def test_codex_agent_config_is_distributed(self):
        manifest = json.loads((ROOT / 'skill-package.json').read_text())
        self.assertIn('agents/*', manifest['include'])
        self.assertTrue((ROOT / 'agents/openai.yaml').is_file())

    def test_catalog_release_matches_package_version(self):
        manifest = json.loads((ROOT / 'skill-package.json').read_text())
        catalog = json.loads((ROOT / 'ai-catalog.json').read_text())
        entry = next(item for item in catalog['entries']
                     if item.get('name') == manifest['name'])
        version = manifest['version']
        self.assertEqual(entry['release_tag'], f'skill-v{version}')
        self.assertEqual(entry['asset'], f"{manifest['name']}-{version}.skill.zip")

    def test_hidden_install_parent_does_not_hide_source(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / '.agent' / 'skill' / 'examples'
            source.mkdir(parents=True)
            content = '---\nid: "1.1"\n---\n# 1.1 Example\n\nFictional fixture.\n'
            (source / '1.1 Example.md').write_text(content)
            hidden = source / '.ignored'
            hidden.mkdir()
            (hidden / '1.2 Hidden.md').write_text('---\nid: "1.2"\n---\n# 1.2 Hidden\n\nmust not be loaded\n\n[[Missing hidden]]\n')   # 按节文件命名：若被读入，必报断链（11-16 反向验证过）
            # 11-16：check_src 的 S9（速查目录按全书结构）与 S11（私有库禁用词表）对样例必然报错，不再断言整体退出码；
            # 只断言本测试关心的：隐藏目录里的文件不被读入
            out = subprocess.run([sys.executable, str(ROOT / 'tools/check_src.py'), '--src', str(source), '--quiet'], capture_output=True, text=True)
            self.assertNotIn('Missing hidden', out.stdout + out.stderr)
            output = Path(tmp) / 'build'
            subprocess.run([sys.executable, str(ROOT / 'tools/assemble.py'), '--src', str(source), '--out', str(output)], check=True, capture_output=True)
            self.assertIn('Fictional fixture.', (output / 'full.md').read_text())

    def test_export_is_current(self):
        subprocess.run([sys.executable, str(ROOT / 'tools/export_prompt.py'), '--check'], check=True)

    def test_source_accepts_valid_and_rejects_broken_reference(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / '1.1 Example.md'
            text = '---\nid: "1.1"\n---\n# 1.1 Example\n\nFictional format fixture.\n'
            path.write_text(text)
            cmd = [sys.executable, str(ROOT / 'tools/check_src.py'), '--src', tmp, '--quiet']
            # 11-16：样例对 S9 / S11 必然报错（见上），改为比对：合法样例没有断链错误，坏引用多出一条断链错误
            ok = subprocess.run(cmd, capture_output=True, text=True).stdout
            self.assertNotIn('断链', ok)
            path.write_text(text + '\n[[Missing note]]\n')
            bad = subprocess.run(cmd, capture_output=True, text=True)
            self.assertNotEqual(bad.returncode, 0)
            self.assertIn('断链 1.1 Example.md -> [[Missing note]]', bad.stdout)

    def test_source_rejects_sd48_forbidden_marks(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / '1.1 Example.md'
            text = '---\nid: "1.1"\n---\n# 1.1 Example\n\nFictional format fixture with 〔待补来源〕.\n'
            path.write_text(text)
            cmd = [sys.executable, str(ROOT / 'tools/check_src.py'), '--src', tmp, '--quiet']
            self.assertNotEqual(subprocess.run(cmd, capture_output=True).returncode, 0)


if __name__ == '__main__':
    unittest.main()
