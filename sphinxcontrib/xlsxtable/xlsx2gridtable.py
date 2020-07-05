import argparse
import unicodedata
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

def parse_index_str(index_str: str) -> int:
    """ Gets index number form Excel column letter (1-based).
        1 from A, 26 from Z and 52 from AZ.
    """
    AZ = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'

    index_str = index_str.upper()
    if index_str[0] in AZ:
        index_str_len = len(index_str)
        index = 0
        for i in range(index_str_len):
            index += (ord(index_str[i]) - ord('A') + 1) * pow(26, index_str_len - 1 - i)
        return index
    else:
        return int(index_str)

def parse_indexes_str(indexes_str: str) -> list:
    """ Gets indexes numbers from string (1-based).
        1, 2, 4 and 6 from '1-2, D-F'.
    """
    indexes = []
    range_indexes_strs = indexes_str.replace(',', ' ').split()
    for range_str in range_indexes_strs:
        if '-' in range_str:
            range_str_els = range_str.split('-')
            begin = parse_index_str(range_str_els[0])
            end = parse_index_str(range_str_els[1])
            indexes += list(range(begin, end + 1))
        else:
            indexes.append(int(range_str))

    return list(set(indexes))

def get_use_indexes(min_index: int, max_index: int, includes_str: str, excludes_str: str):
    """ Gets valid indexes (1-based).
    """
    includes = set(range(min_index, max_index + 1))
    if includes_str is not None and len(includes_str) > 0:
        includes = set(parse_indexes_str(includes_str))

    excludes = set([])
    if excludes_str is not None and len(excludes_str) > 0:
        excludes = set(parse_indexes_str(excludes_str))

    return list(includes - excludes)

def get_string_width(text: str):
    width = 0
    for t in text:
        if unicodedata.east_asian_width(t) in 'FW':
            width += 2
        else:
            width += 1
    return width

def get_max_width(lines):
    max_w = 0
    for l in lines:
        max_w = max(max_w, get_string_width(l))
    return max_w

def count_line(values):
    count = len(values)
    if '|' in ''.join(values):
        count += 1
    return count

class TableCell:
    is_merged_top = False
    is_merged_left = False

    def __init__(self, row, column, value):
        values = []
        if value != None:
            values = f'{value}'.split('\n')

        self.row = row
        self.column = column
        self.values = values
        self.line_count = count_line(values)
        self.width = get_max_width(values)

def get_padding(count, max_count):
    padding = ''
    for _ in range(max_count - count):
        padding += ' '
    return padding

def get_rule(colmuns, is_head=False, is_end=False):
    line_str = ''
    for cell in colmuns:
        line_str += '+'
        for _ in range(cell.width + 2):
            if cell.is_merged_top and not is_end:
                line_str += ' '
            elif is_head:
                line_str += '='
            else:
                line_str += '-'

    line_str += '+'
    return line_str

def gen_reST_grid_table_lines(
    filename: str,
    header_rows=0,
    sheetname=None,
    start_row=1,
    start_column=1,
    include_rows=None,
    exclude_rows=None,
    include_columns=None,
    exclude_columns=None):

    wb = load_workbook(
        filename=filename,
        read_only=False, # Can not get merged cell information if read_only is True
        keep_vba=False,
        data_only=True,
        keep_links=False
        )

    # select target sheet
    try:
        ws = wb[sheetname]
    except:
        ws = wb.active

    # rows / columns
    offset_row = max(ws.min_row, start_row)
    offset_col = max(ws.min_column, start_column)

    use_rows = get_use_indexes(ws.min_row, ws.max_row, include_rows, exclude_rows)
    use_cols = get_use_indexes(ws.min_column, ws.max_column, include_columns, exclude_columns)

    # parse cell info
    table_cells = []
    for r in range(offset_row, ws.max_row + 1):
        if r in use_rows:
            # appebd array for row
            table_cells.append([])

            # get line count in the row
            r_index = len(table_cells) - 1
            for c  in range(offset_col, ws.max_column + 1):
                if c in use_cols:
                    tc = TableCell(r, c, ws.cell(r, c).value)
                    table_cells[r_index].append(tc)

    # adjust line count
    row_count = len(table_cells)
    col_count = len(table_cells[0])
    for r in range(row_count):
        max_line_count = 0
        for c in range(col_count):
            max_line_count = max(max_line_count, table_cells[r][c].line_count)

        for c in range(col_count):
            table_cells[r][c].line_count = max_line_count

    # adjust width
    for c in range(col_count):
        max_width = 0
        for r in range(row_count):
            max_width = max(max_width, table_cells[r][c].width)

        for r in range(row_count):
            table_cells[r][c].width = max_width

    # adjust merged cell info
    for mrange in ws.merged_cell_ranges:
        left = mrange.bounds[0] - 1
        top = mrange.bounds[1] - 1
        right = mrange.bounds[2] - 1
        bottom = mrange.bounds[3] - 1
        for c in range(left, right + 1):
            for r in range(top, bottom + 1):
                table_cells[r][c].is_merged_top = (r != top)
                table_cells[r][c].is_merged_left = (c != left)

    # gen lines
    grid_table_lines = []
    for r in range(0, len(table_cells)):
        cols = table_cells[r]
        if r == header_rows and header_rows > 0:
            grid_table_lines.append(get_rule(cols, True))
        else:
            grid_table_lines.append(get_rule(cols))

        line_count = cols[0].line_count
        for l in range(line_count):
            line_str = ''
            for cell in cols:
                line_str += '| ' if cell.is_merged_left == False else '  '
                if len(cell.values) > l:
                    line_str += f'{cell.values[l]}'
                    line_str += get_padding(get_string_width(cell.values[l]), cell.width)
                    line_str += ' '
                else:
                    line_str += get_padding(0, cell.width)
                    line_str += ' '
            line_str += '|'

            grid_table_lines.append(line_str)

    grid_table_lines.append(get_rule(cols, is_end=True))
    return grid_table_lines

def draw_reST_grid_table(
        filename, header_rows, sheet, start_row, start_column,
        include_rows, exclude_rows,
        include_columns, exclude_columns):
    lines = gen_reST_grid_table_lines(filename, header_rows, sheet, start_row, start_column,
        include_rows, exclude_rows, include_columns, exclude_columns)
    for l in lines:
        print(l)

def main():
    p = argparse.ArgumentParser(description='Grid Table String Generator')
    p.add_argument('--header-rows', type=int, default=0, help='Header rows')
    p.add_argument('--sheet', type=str, help='Target sheet name')
    p.add_argument('--start-row', type=int, default=1, help='Start row')
    p.add_argument('--start-column', type=int, default=1, help='Start colmun')
    p.add_argument('--include-rows', type=str, default=None, help='Specify included rows')
    p.add_argument('--exclude-rows', type=str, default=None, help='Specify excluded rows')
    p.add_argument('--include-columns', type=str, default=None, help='Specify included columns')
    p.add_argument('--exclude-columns', type=str, default=None, help='Specify excluded columns')
    p.add_argument('file', type=str, help='Target Excel file path')

    args = p.parse_args()

    draw_reST_grid_table(args.file, args.header_rows, args.sheet,
        args.start_row, args.start_column,
        args.include_rows, args.exclude_rows,
        args.include_columns, args.exclude_columns)

if __name__ == '__main__':
    main()
