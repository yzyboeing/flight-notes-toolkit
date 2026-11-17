import importlib.util
from pathlib import Path
import unittest


module_path = Path(__file__).resolve().parents[1] / 'tools/check_src.py'
spec = importlib.util.spec_from_file_location('check_src', module_path)
checker = importlib.util.module_from_spec(spec)
spec.loader.exec_module(checker)


class SharedPowerDifferenceTests(unittest.TestCase):
    def setUp(self):
        self.steps = '多做三步：<br>示例动作一；<br>示例电门 OFF；<br>示例设备 OFF'
        self.summary = ('<table><tr><td><strong>仅剩一个主交流电源后</strong><br>（APU 不可用）</td>'
                        '<td>计划在最近合适机场着陆；《示例甲》《示例乙》《示例丙》检查单结束，其余检查单直接进入下一步。</td>'
                        '<td>计划在最近合适机场着陆，然后' + self.steps + '</td></tr></table>')
        self.cells = ['仅剩一个主交流电源后', '检查单结束', self.steps]

    def test_explicitly_listed_checklists_match_both_complete_cells(self):
        for title in ['示例甲', '示例乙', '示例丙']:
            self.assertTrue(checker.shared_power_difference_covered(self.summary, title, self.cells))

    def test_removed_step_is_reported(self):
        cells = self.cells.copy()
        cells[2] = self.steps.replace('<br>示例设备 OFF', '')
        self.assertFalse(checker.shared_power_difference_covered(self.summary, '示例乙', cells))

    def test_changed_switch_position_is_reported(self):
        cells = self.cells.copy()
        cells[2] = self.steps.replace('OFF', 'ON', 1)
        self.assertFalse(checker.shared_power_difference_covered(self.summary, '示例乙', cells))

    def test_checklist_removed_from_shared_applicability_is_reported(self):
        summary = self.summary.replace('《示例乙》', '')
        self.assertFalse(checker.shared_power_difference_covered(summary, '示例乙', self.cells))

    def test_changed_number_and_qualifier_are_reported(self):
        for index, value in [(0, '仅剩两个主交流电源后'), (1, '检查单结束或直接进入下一步')]:
            cells = self.cells.copy()
            cells[index] = value
            self.assertFalse(checker.shared_power_difference_covered(self.summary, '示例乙', cells))

    def test_unrelated_or_malformed_summary_is_reported(self):
        self.assertFalse(checker.shared_power_difference_covered(self.summary, '电气', self.cells))
        self.assertFalse(checker.shared_power_difference_covered(self.summary.replace('然后', '随后'), '示例乙', self.cells))


if __name__ == '__main__':
    unittest.main()
