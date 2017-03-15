# coding: utf-8
from __future__ import (
    absolute_import,
    print_function,
    unicode_literals,
)

from pydocx.models import XmlModel, XmlChild


class TableCellProperties(XmlModel):
    XML_TAG = 'tcPr'

    grid_span = XmlChild(name='gridSpan', attrname='val')

    background_fill = XmlChild(name='shd', attrname='fill')

    vertical_merge = XmlChild(name='vMerge', type=lambda el: dict(el.attrib))  # noqa

    def should_close_previous_vertical_merge(self):
        # If vMerge is omitted, then this cell shall not be part of any
        # vertically merged grouping of cells, and any vertically merged group
        # of preceding cells shall be closed.
        if self.vertical_merge is None:
            return True
        return not self.is_continue_vertical_merge

    @property
    def is_continue_vertical_merge(self):
        if self.vertical_merge is None:
            return False
        merge = self.vertical_merge.get('val', 'continue')
        return merge == 'continue'

    @property
    def background_color(self):
        background_fill = getattr(self, 'background_fill', None)

        # There is no need to set white background color
        if background_fill not in ('auto', 'FFFFFF'):
            return background_fill

        return None
