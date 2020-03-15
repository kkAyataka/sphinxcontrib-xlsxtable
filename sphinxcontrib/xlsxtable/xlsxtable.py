import os
from docutils import nodes
from docutils.parsers.rst import directives
from docutils.statemachine import ViewList

from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

from .xlsx2gridtable import gen_reST_grid_table_lines

class XlsxTable(directives.tables.RSTTable):
    has_content = True

    optional_arguments = 1
    option_spec = {
        'file': directives.path,
        'header-rows': directives.positive_int,
        'sheet': directives.unchanged,
    }

    def run(self):
        filepath = self.options.get('file', '')
        header_rows = self.options.get('header-rows', 0)
        sheet = self.options.get('sheet', None)

        rst_dir = os.path.dirname(os.path.abspath(self.state.document.current_source))
        filepath = os.path.normpath(os.path.join(rst_dir, filepath))

        lines = gen_reST_grid_table_lines(filepath, header_rows, sheet)
        node = nodes.Element(rawsource='\n'.join(lines))

        #for l in lines:
        #    print(l)

        title, messages = self.make_title()

        self.content = ViewList(lines, self.content.source)
        self.state.nested_parse(self.content, self.content_offset, node)
        table_node = node[0]
        if title:
            table_node.insert(0, title)

        return [table_node] + messages

def setup(app):
    app.add_directive("xlsx-table", XlsxTable)

    return {
        'version': '0.1.9',
        'parallel_read_safe': True,
        'parallel_write_safe': True,
    }
