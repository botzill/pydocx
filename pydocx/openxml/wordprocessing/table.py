# coding: utf-8
from __future__ import (
    absolute_import,
    print_function,
    unicode_literals,
)

from collections import defaultdict

from pydocx.models import XmlModel, XmlCollection, XmlChild
from pydocx.openxml.wordprocessing.table_row import TableRow
from pydocx.openxml.wordprocessing.table_properties import TableProperties


class Table(XmlModel):
    XML_TAG = 'tbl'

    properties = XmlChild(type=TableProperties)

    rows = XmlCollection(
        TableRow,
    )

    def calculate_table_cell_spans(self):
        if not self.rows:
            return

        active_rowspan_cells_by_column = {}
        cell_to_rowspan_count = defaultdict(int)
        for row in self.rows:
            for column_index, cell in enumerate(row.cells):
                properties = cell.properties
                # If this element is omitted, then this cell shall not be
                # part of any vertically merged grouping of cells, and any
                # vertically merged group of preceding cells shall be
                # closed.
                if properties is None or properties.vertical_merge is None:
                    # if properties are missing, this is the same as the
                    # the element being omitted
                    active_rowspan_cells_by_column[column_index] = None
                elif properties:
                    vertical_merge = properties.vertical_merge.get('val', 'continue')  # noqa
                    if vertical_merge == 'restart':
                        active_rowspan_cells_by_column[column_index] = cell
                        cell_to_rowspan_count[cell] += 1
                    elif vertical_merge == 'continue':
                        active_rowspan_for_column = active_rowspan_cells_by_column.get(column_index)  # noqa
                        if active_rowspan_for_column:
                            cell_to_rowspan_count[active_rowspan_for_column] += 1  # noqa
        return dict(cell_to_rowspan_count)

    def get_style_chain_stack(self):
        if not self.properties:
            return

        parent_style = self.properties.parent_style
        if not parent_style:
            return

        part = getattr(self.container, 'style_definitions_part', None)
        if part:
            style_stack = part.get_style_chain_stack('table', parent_style)
            for result in style_stack:
                yield result

    def get_paragraph_properties(self):
        """Get default style paragraph properties for table"""

        paragraph_properties = None
        for style in self.get_style_chain_stack():
            if style.paragraph_properties:
                paragraph_properties = style.paragraph_properties
                break

        return paragraph_properties
