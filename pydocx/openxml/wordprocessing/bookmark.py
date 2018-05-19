# coding: utf-8
from __future__ import (
    absolute_import,
    print_function,
    unicode_literals,
)

from pydocx.models import XmlModel, XmlAttribute
from pydocx.util.memoize import memoized


class Bookmark(XmlModel):
    XML_TAG = 'bookmarkStart'

    name = XmlAttribute(name='name')

    @memoized
    def get_name(self):
        name = self.name
        if name:
            # The _Goback bookmark enables to go to the previous edit functionality (Shift+F5) across sessions.
            # It is a hidden bookmark
            if name.lower() in ('_goback',):
                return None
        return name
