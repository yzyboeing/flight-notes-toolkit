import importlib.util
from pathlib import Path
import unittest

spec = importlib.util.spec_from_file_location('note_checker', Path(__file__).parents[1] / 'tools/check_src.py')
m = importlib.util.module_from_spec(spec)
spec.loader.exec_module(m)

class NoteHints(unittest.TestCase):
    def test_counts_are_per_entry_and_four_is_threshold(self):
        text = '### 1. Alpha\n' + '注：标签——正文\n' * 3 + '### 2. Beta\n' + '注：标签——正文\n' * 4
        count, empty, meta = m.check_note_lines(text, 'sample.md')
        self.assertEqual(len(count), 1)
        self.assertIn('2. Beta', count[0])
        self.assertEqual(empty + meta, [])

    def test_empty_after_removing_markup_and_absolute_line(self):
        hints = m.check_note_lines('### 1. Alpha\n注：<br>  \n注：标签——正文\n', 'sample.md', offset=7)
        self.assertEqual(hints[1], ['sample.md:9'])

    def test_meta_keywords_do_not_flag_normal_operating_text(self):
        text = '### 1. Alpha\n注：译法——中文 QRH 译作示例\n注：范围——本表以示例为准\n注：动作——执行《示例》检查单。\n'
        hints = m.check_note_lines(text, 'sample.md')
        self.assertEqual(len(hints[2]), 2)

    def test_all_three_hints_can_coexist(self):
        text = '### 1. Alpha\n注：\n注：译法——属翻译差异\n注：甲——正文\n注：乙——正文\n'
        self.assertEqual([len(v) for v in m.check_note_lines(text, 'sample.md')], [1, 1, 1])

if __name__ == '__main__': unittest.main()
