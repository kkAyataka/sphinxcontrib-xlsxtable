import os
from docutils import nodes
from docutils.parsers.rst import directives
from docutils.statemachine import ViewList

from .xlsx2gridtable import gen_reST_grid_table_lines

class XlsxTable(directives.tables.RSTTable):
    has_content = True

    optional_arguments = 1
    option_spec = {
        'file': directives.path,
        'header-rows': directives.positive_int,
        'start-row': directives.positive_int,
        'start-column': directives.positive_int,
        'include-rows': directives.unchanged,
        'exclude-rows': directives.unchanged,
        'include-columns': directives.unchanged,
        'exclude-columns': directives.unchanged,
        'sheet': directives.unchanged,
    }

    def run(self):
        file = self.options.get('file', '')
        header_rows = self.options.get('header-rows', 0)
        sheet = self.options.get('sheet', None)
        start_row = self.options.get('start-row', 1)
        start_column = self.options.get('start-column', 1)
        include_rows = self.options.get('include-rows', None)
        exclude_rows = self.options.get('exclude-rows', None)
        include_columns = self.options.get('include-columns', None)
        exclude_columns = self.options.get('exclude-columns', None)

        rst_dir = os.path.dirname(os.path.abspath(self.state.document.current_source))
        fullpath = os.path.normpath(os.path.join(rst_dir, file))

        lines = gen_reST_grid_table_lines(
            file,
            fullpath,
            header_rows,
            sheet,
            start_row,
            start_column,
            include_rows, exclude_rows,
            include_columns, exclude_columns)
        node = nodes.Element(rawsource='\n'.join(lines))

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
        'version': '1.1.1',
        'parallel_read_safe': True,
        'parallel_write_safe': True,
    }
