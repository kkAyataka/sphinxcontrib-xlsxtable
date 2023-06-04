import os
import textwrap
import unittest
from sphinxcontrib.xlsxtable import xlsx2gridtable

def gen_ok_text(text: str) -> str:
    return textwrap.dedent(text).split('\n')[1:-1]

class TestIndexesParser(unittest.TestCase):
    def test_parse_range(self):
        indexes = xlsx2gridtable.parse_indexes_str('3-10')
        self.assertEqual(indexes, [3, 4, 5, 6, 7, 8, 9, 10])

    def test_parse_range_az(self):
        indexes = xlsx2gridtable.parse_indexes_str('b-z')
        self.assertEqual(indexes, list(range(2, 27)))

    def test_parse_sep_space_comma(self):
        indexes = xlsx2gridtable.parse_indexes_str('A-B AA-BU')
        self.assertEqual(indexes, [1, 2] + list(range(27, 74)))

        indexes = xlsx2gridtable.parse_indexes_str('C-G Z-AB, F-J')
        self.assertEqual(indexes, list(range(3, 11)) + [26, 27, 28])

class TestGenGridTable(unittest.TestCase):
    def test_gen_normal(self):
        xlsxfile = os.path.abspath(f'{os.path.dirname(__file__)}/_res/sample.xlsx')
        res = xlsx2gridtable.gen_reST_grid_table_lines(
            file='./_res/sample.xlsx',
            fullpath=xlsxfile,
            header_rows=0,
            sheetname='Sheet4',
            start_row=0,
            start_column=0,
            include_rows=None,
            exclude_rows=None,
            include_columns=None,
            exclude_columns=None
            )

        ok = gen_ok_text('''
            +-------+----+----+-------+
            | A1    | B1 | C1 | D1    |
            +-------+----+----+-------+
            | A2:B2      | C2 | D2:D3 |
            +-------+----+----+       +
            | A3    | B3 | C3 |       |
            +-------+----+----+-------+
            | A4    | B4 | C4 | D4    |
            +-------+----+----+-------+
            | A5:B6      | C5 | D5    |
            +            +----+-------+
            |            | C6 | D6    |
            +-------+----+----+-------+
            | A7    | B7 | C7 | D7    |
            +-------+----+----+-------+
            ''')

        self.assertListEqual(res, ok)

    def test_head_start_row_start_column(self):
        xlsxfile = os.path.abspath(f'{os.path.dirname(__file__)}/_res/sample.xlsx')
        res = xlsx2gridtable.gen_reST_grid_table_lines(
            file='./_res/sample.xlsx',
            fullpath=xlsxfile,
            header_rows=1,
            sheetname='Sheet4',
            start_row=4,
            start_column=3,
            include_rows=None,
            exclude_rows=None,
            include_columns=None,
            exclude_columns=None
            )

        ok = gen_ok_text('''
            +----+----+
            | C4 | D4 |
            +====+====+
            | C5 | D5 |
            +----+----+
            | C6 | D6 |
            +----+----+
            | C7 | D7 |
            +----+----+
            ''')

        self.assertListEqual(res, ok)

    def test_include_rows_exclude_rows_include_columns_exclude_columns(self):
        xlsxfile = os.path.abspath(f'{os.path.dirname(__file__)}/_res/sample.xlsx')
        res = xlsx2gridtable.gen_reST_grid_table_lines(
            file='./_res/sample.xlsx',
            fullpath=xlsxfile,
            header_rows=0,
            sheetname='Sheet4',
            start_row=0,
            start_column=0,
            include_rows='2-6',
            exclude_rows='1-2 4-100',
            include_columns='B-C',
            exclude_columns='A B D'
            )

        ok = gen_ok_text('''
            +----+
            | C3 |
            +----+
            ''')

        self.assertListEqual(res, ok)

    def test_divide_merged_cell(self):
        xlsxfile = os.path.abspath(f'{os.path.dirname(__file__)}/_res/sample.xlsx')
        res = xlsx2gridtable.gen_reST_grid_table_lines(
            file='./_res/sample.xlsx',
            fullpath=xlsxfile,
            header_rows=1,
            sheetname='Sheet4',
            start_row=0,
            start_column=0,
            include_rows=None,
            exclude_rows='1-2',
            include_columns=None,
            exclude_columns='A'
            )

        ok = gen_ok_text('''
            +----+----+----+
            | B3 | C3 |    |
            +====+====+====+
            | B4 | C4 | D4 |
            +----+----+----+
            |    | C5 | D5 |
            +    +----+----+
            |    | C6 | D6 |
            +----+----+----+
            | B7 | C7 | D7 |
            +----+----+----+
            ''')

        self.assertListEqual(res, ok)

    def test_images(self):
        xlsxfile = os.path.abspath(f'{os.path.dirname(__file__)}/_res/sample-images.xlsx')
        res = xlsx2gridtable.gen_reST_grid_table_lines(
            file='./_res/sample-images.xlsx',
            fullpath=xlsxfile,
            header_rows=0,
            sheetname='v',
            start_row=0,
            start_column=0,
            include_rows=None,
            exclude_rows=None,
            include_columns=None,
            exclude_columns=None
            )

        ok = gen_ok_text('''
            +-------+-------------------------------------------------------+
            | Red   | .. image:: ./_res/sample-images.xlsx.media/image1.png |
            +-------+-------------------------------------------------------+
            | Green | .. image:: ./_res/sample-images.xlsx.media/image2.png |
            +-------+-------------------------------------------------------+
            | Blue  | .. image:: ./_res/sample-images.xlsx.media/image3.png |
            +-------+-------------------------------------------------------+
            ''')

        self.assertListEqual(res, ok)

if __name__ == '__main__':
    unittest.main()
