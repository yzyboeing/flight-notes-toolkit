import importlib.util
from pathlib import Path
import unittest
p = Path(__file__).parents[1] / 'tools' / 'check_src.py'
s = importlib.util.spec_from_file_location('checker', p)
m = importlib.util.module_from_spec(s); s.loader.exec_module(m)

class TypeMixtures(unittest.TestCase):
    def test_cell_and_multiline_paragraph(self):
        text = '<table><tr><td>【737-NG】alpha\n【737-8】beta</td></tr></table>\n\n段【737-NG】alpha\n续【737-8】beta\n'
        hits = m.find_type_mixtures(text)
        self.assertEqual([x['kind'] for x in hits], ['cell', 'paragraph'])
        self.assertEqual([x['line'] for x in hits], [1, 4])
    def test_separate_cells_and_paragraphs(self):
        text = '<table><tr><td>【737-NG】a</td><td>【737-8】b</td></tr></table>\n\n【737-NG】a\n\n【737-8】b'
        self.assertEqual(m.find_type_mixtures(text), [])
    def test_repeated_cells_keep_individual_offsets(self):
        cell = '<td>【737-NG】a【737-8】b</td>'
        hits = m.find_type_mixtures('<table>\n<tr>'+cell+'</tr>\n<tr>'+cell+'</tr>\n</table>')
        self.assertEqual([x['line'] for x in hits], [2, 3])
    def test_comments_and_incomplete_markers(self):
        self.assertEqual(m.find_type_mixtures('<!-- 【737-NG】a【737-8】b -->\n\n【737-NG a【737-8】b'), [])

if __name__ == '__main__': unittest.main()
