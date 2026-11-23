"""SD-275：同一网格列句末标点统一；K-S43：表格各行实占列数一致（合成样例，不含笔记正文）。"""
import importlib.util
from pathlib import Path
import subprocess, sys, tempfile, unittest

TOOLS = Path(__file__).parents[1] / 'tools'
spec = importlib.util.spec_from_file_location('check_src_mod', TOOLS / 'check_src.py')
cs = importlib.util.module_from_spec(spec); spec.loader.exec_module(cs)

TABLE = '''### 1. 样例

<table class="ftn">
<tr class="hdr"><th>项目</th><th>甲</th><th colspan="2">乙</th></tr>
<tr><td>一</td><td>系统接通后保持当前航迹。</td><td colspan="2">按压电门后系统转到备用方式。</td></tr>
<tr><td rowspan="3">二</td><td rowspan="3">系统在地面时自动转到备用方式。</td><td class="center"><strong>时机</strong></td><td class="center"><strong>速度</strong></td></tr>
<tr><td>之前</td><td>所选速度</td></tr>
<tr><td>之后</td><td>所选进近速度。</td></tr>
<tr><td>三</td><td>空白</td><td colspan="2">选择另一方式后指令杆保持当前姿态</td></tr>
<tr><td>四</td><td>备用数据源</td><td colspan="2">系统转到人工方式。</td></tr>
</table>
'''


def run_punct(text):
    with tempfile.TemporaryDirectory() as d:
        f = Path(d) / 'notes_src' / '1 样例' / '1.1 样例.md'
        f.parent.mkdir(parents=True)
        f.write_text(text, encoding='utf-8')
        subprocess.run([sys.executable, str(TOOLS / 'punct_rule.py'), '--repo', d, '--apply'], check=True, capture_output=True)
        out = f.read_text(encoding='utf-8')
        again = subprocess.run([sys.executable, str(TOOLS / 'punct_rule.py'), '--repo', d], check=True, capture_output=True, text=True).stdout
        return out, again


class PunctColumns(unittest.TestCase):
    def test_phrase_column_drops_stray_period(self):
        out, _ = run_punct(TABLE)
        self.assertIn('<td>所选进近速度</td>', out)          # 同列全是短语：多出来的句号去掉
        self.assertIn('<strong>速度</strong>', out)           # 表身小表头不动

    def test_sentence_column_and_state_word(self):
        out, _ = run_punct(TABLE)
        self.assertIn('指令杆保持当前姿态。</td>', out)      # 同列多数是句子：补句号
        self.assertIn('<td>空白</td>', out)                  # 状态词永不加句号
        self.assertIn('<td>备用数据源。</td>', out)   # 短语在多数是句子的列里随列补句号（旧规则 A 逐格判会漏）

    def test_idempotent(self):
        _, again = run_punct(TABLE)
        self.assertIn('高把握改 0', again)

    def test_row_occupancy(self):
        good = TABLE.split('\n', 2)[2]
        self.assertEqual(len(set(cs.table_occupancy(good))), 1)
        bad = good.replace('<td colspan="2">选择另一方式', '<td>选择另一方式')
        occ = cs.table_occupancy(bad)
        self.assertEqual(min(occ), 3)
        self.assertEqual(max(occ), 4)


if __name__ == '__main__':
    unittest.main()
