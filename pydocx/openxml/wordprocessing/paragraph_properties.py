# coding: utf-8
from __future__ import (
    absolute_import,
    print_function,
    unicode_literals,
)

from pydocx.models import XmlModel, XmlChild
from pydocx.openxml.wordprocessing.numbering_properties import NumberingProperties  # noqa
from pydocx.types import OnOff


class ParagraphProperties(XmlModel):
    XML_TAG = 'pPr'

    parent_style = XmlChild(name='pStyle', attrname='val')
    numbering_properties = XmlChild(type=NumberingProperties)
    justification = XmlChild(name='jc', attrname='val')
    # TODO ind can appear multiple times. Need to merge them in document order
    # This probably means other elements can appear multiple times

    # TODO Left/right is for traditional conformance. Need to handle start/end
    # for strict conformance
    indentation_left = XmlChild(name='ind', attrname='left')
    indentation_right = XmlChild(name='ind', attrname='right')
    indentation_first_line = XmlChild(name='ind', attrname='firstLine')
    indentation_hanging = XmlChild(name='ind', attrname='hanging')

    # paragraph spacing
    spacing_after = XmlChild(name='spacing', attrname='after')
    spacing_line = XmlChild(name='spacing', attrname='line')
    spacing_line_rule = XmlChild(name='spacing', attrname='lineRule')
    spacing_after_auto_spacing = XmlChild(type=OnOff, name='spacing', attrname='afterAutospacing')

    @property
    def start_margin_position(self):
        # Regarding indentation,
        #   position = left - hanging
        #   position = left + firstLine (only if hanging isn't specified)
        # 17.3.1.12 - The firstLine and hanging attributes are mutually
        # exclusive, if both are specified, then the firstLine value is
        # ignored.
        start_margin = 0
        if self.indentation_left:
            start_margin += int(float(self.indentation_left))
        if self.indentation_hanging:
            start_margin -= int(float(self.indentation_hanging))
        elif self.indentation_first_line:
            start_margin += int(float(self.indentation_first_line))
        if start_margin:
            return start_margin
        return 0

    def to_int(self, attribute, default=None):
        # TODO would be nice if this integer conversion was handled
        # implicitly by the model somehow
        try:
            return int(getattr(self, attribute, default))
        except (ValueError, TypeError):
            return default

    @property
    def is_list_paragraph(self):
        return self.parent_style == 'ListParagraph'

    @property
    def no_indentation(self):
        return not any((
            self.indentation_left,
            self.indentation_hanging,
            self.indentation_right,
            self.indentation_first_line,
        ))

    @property
    def no_spacing(self):
        return not any((
            self.spacing_line,
            self.spacing_after,
            self.spacing_after_auto_spacing,
            self.spacing_line_rule,
        ))
