from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]


class PackageTest(unittest.TestCase):
    def test_hidden_install_parent_does_not_hide_source(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / '.agent' / 'skill' / 'examples'
            source.mkdir(parents=True)
            content = '---\nid: "1.1"\n---\n# 1.1 Example\n\nFictional fixture.\n'
            (source / '1.1 Example.md').write_text(content)
            hidden = source / '.ignored'
            hidden.mkdir()
            (hidden / 'invalid.md').write_text('must not be loaded')
            subprocess.run([sys.executable, str(ROOT / 'tools/check_src.py'), '--src', str(source), '--quiet'], check=True, capture_output=True)
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
            self.assertEqual(subprocess.run(cmd, capture_output=True).returncode, 0)
            path.write_text(text + '\n[[Missing note]]\n')
            self.assertNotEqual(subprocess.run(cmd, capture_output=True).returncode, 0)


if __name__ == '__main__':
    unittest.main()
