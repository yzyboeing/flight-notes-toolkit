import importlib.util
from pathlib import Path
import unittest

spec = importlib.util.spec_from_file_location('check_src', Path(__file__).resolve().parents[1] / 'tools/check_src.py')
checker = importlib.util.module_from_spec(spec)
spec.loader.exec_module(checker)

class EditorialMarkers(unittest.TestCase):
    def test_plain_and_html_labels_keep_line_numbers(self):
        source = '<!-- 解释：内部记录\n继续 -->\n公司差异：要点\n<tr><td>正常<br><small>解释：</small>旧写法</td></tr>'
        retired, named, abbr = checker.scan_editorial_markers(source, '速查/速查源.md', set())
        self.assertEqual(retired, ['速查/速查源.md:3「公司差异：」','速查/速查源.md:4「解释：」'])
        self.assertEqual((named, abbr), ([], []))

    def test_named_manual_prose_outside_note(self):
        rows = checker.scan_editorial_markers('SOP 写作某动作\n运行手册只写某条件\n公司手册规定……', '正文.md', set(), 20)
        self.assertEqual(rows[1], ['正文.md:21「SOP 写作」','正文.md:22「运行手册只写」','正文.md:23「公司手册」'])

    def test_additional_named_manual_phrases(self):
        source = 'FCOM 限制\nFCOM 用 1600m\n大纲用 1.5km\nQRH：步骤'
        rows = checker.scan_editorial_markers(source, '正文.md', set())
        self.assertEqual(rows[1], ['正文.md:1「FCOM 限制」', '正文.md:2「FCOM 用」',
                                   '正文.md:3「大纲用」', '正文.md:4「QRH：」'])

    def test_shared_model_columns_respect_rowspans(self):
        source = '''<table>
<tr><th rowspan="2">项目</th><th colspan="2">机型</th></tr>
<tr><th>737-NG</th><th>737-8</th></tr>
<tr><td rowspan="2">程序</td><td colspan="2">共有步骤</td></tr>
<tr><td colspan="2">仅限 <b>【737-8】</b></td></tr>
<tr><td colspan="2">部分 737-NG 飞机</td><td>允许</td></tr>
</table>'''
        self.assertEqual(checker.scan_model_colspans(source, '正文.md', 10),
                         ['正文.md:15「【737-8】」'])

    def test_shared_model_columns_flag_unbracketed_marker(self):
        source = '''<table>
<tr><th>项目</th><th>737-NG</th><th>737-8</th></tr>
<tr><td>应答机</td><td colspan="2">737-8 及部分 737-NG 飞机</td></tr>
</table>'''
        self.assertEqual(checker.scan_model_colspans(source, '速查源.md'),
                         ['速查源.md:3「737-8 及」'])

    def test_manual_koujing_and_generic_phrase(self):
        source = 'QRH 口径：调定所需位置\nSOP 的口径：核实条件\n运行口径：独立分组'
        rows = checker.scan_editorial_markers(source, '正文.md', set())
        self.assertEqual(rows[1], ['正文.md:1「QRH 口径」', '正文.md:2「SOP 的口径」'])

    def test_common_abbreviation_only_in_glossary(self):
        common = checker.common_abbreviations()
        self.assertIn('V/S', common)
        self.assertIn('FMC', common)
        source = '原理解释：正常语句\nFMC：系统说明\n注：缩写——FMC：名称；LOUT：名称；V/S：名称'
        rows = checker.scan_editorial_markers(source, '正文.md', common)
        self.assertEqual(rows, ([], [], ['正文.md:3「FMC」','正文.md:3「V/S」']))

if __name__ == '__main__':
    unittest.main()
