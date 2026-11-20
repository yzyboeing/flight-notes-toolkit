"""2026-11-20 规则瘦身：K-S39～S42 数字规则与 check_rules 决策记录 / 体量上限的正反测试。"""
import importlib.util
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location('check_src', ROOT / 'tools/check_src.py')
check_src = importlib.util.module_from_spec(spec)
spec.loader.exec_module(check_src)

TABLE = '<table class="ftn">\n<tr><td>甲</td><td>乙</td></tr>\n</table>'


class NumericRules(unittest.TestCase):
    def run_rules(self, body):
        return check_src.check_numeric_rules(body, 'x.md')

    def test_five_notes_is_error_four_is_not(self):
        four = '### 1. 标题\n\n' + '\n'.join('注：标签——内容%d。' % i for i in range(4)) + '\n'
        five = '### 1. 标题\n\n' + '\n'.join('注：标签——内容%d。' % i for i in range(5)) + '\n'
        self.assertEqual(self.run_rules(four)[0], [])
        self.assertEqual(len(self.run_rules(five)[0]), 1)

    def test_back_to_back_tables_need_caption(self):
        bad = '### 1. 标题\n\n%s\n\n%s\n' % (TABLE, TABLE)
        good = '### 1. 标题\n\n%s\n\n两型差异：\n\n%s\n' % (TABLE, TABLE)
        pair = '### 1. 标题\n\n%%%%TABPAIR%%%% 50\n%s\n\n%s\n' % (TABLE, TABLE)
        self.assertEqual(len(self.run_rules(bad)[1]), 1)
        self.assertEqual(self.run_rules(good)[1], [])
        self.assertEqual(self.run_rules(pair)[1], [])

    def test_caption_above_callouts(self):
        bad = '### 1. 标题\n\n737-8：\n\n<!-- callouts:x -->\n<!-- /callouts:x -->\n'
        skeleton = '### 1. 标题\n\n检查单与重要项目说明：\n\n<!-- callouts:x -->\n<!-- /callouts:x -->\n'
        self.assertEqual(len(self.run_rules(bad)[2]), 1)
        self.assertEqual(self.run_rules(skeleton)[2], [])

    def test_mechanically_merged_note(self):
        bad = '### 1. 标题\n\n注：补充——<br>甲。<br>乙。\n'
        good = '### 1. 标题\n\n注：补充——甲。\n'
        self.assertEqual(len(self.run_rules(bad)[3]), 1)
        self.assertEqual(self.run_rules(good)[3], [])


class RulesMaintenance(unittest.TestCase):
    def run_check(self, rules_text, decisions_text):
        with tempfile.TemporaryDirectory() as tmp:
            notes = Path(tmp)
            (notes / '规则').mkdir()
            (notes / '规则/现行规则.md').write_text(rules_text, encoding='utf-8')
            (notes / '规则/决策记录.md').write_text(decisions_text, encoding='utf-8')
            return subprocess.run([sys.executable, str(ROOT / 'tools/check_rules.py'), str(notes)],
                                  capture_output=True, text=True)

    def test_decision_without_landing_line_fails(self):
        result = self.run_check('# 规则\n', '# 决策记录\n\n## SD-267　新决定（2026-11-21，用户）\n- 定下：……\n')
        self.assertNotEqual(result.returncode, 0)
        self.assertIn('缺「落实：」行', result.stdout)

    def test_decision_with_landing_line_passes(self):
        result = self.run_check('# 规则\n', '# 决策记录\n\n## SD-267　新决定（2026-11-21，用户）\n- 落实：现行规则 第 3 节（3.5）\n')
        self.assertEqual(result.returncode, 0, result.stdout)

    def test_rules_over_size_limit_fails(self):
        result = self.run_check('# 规则\n' + '字' * 11000, '# 决策记录\n')
        self.assertNotEqual(result.returncode, 0)
        self.assertIn('超过上限', result.stdout)


if __name__ == '__main__':
    unittest.main()
