# coding: utf-8
from __future__ import (
    absolute_import,
    print_function,
    unicode_literals,
)

from pydocx.models import XmlModel, XmlCollection, XmlChild
from pydocx.openxml.wordprocessing.bookmark import Bookmark
from pydocx.openxml.wordprocessing.br import Break
from pydocx.openxml.wordprocessing.deleted_run import DeletedRun
from pydocx.openxml.wordprocessing.hyperlink import Hyperlink
from pydocx.openxml.wordprocessing.inserted_run import InsertedRun
from pydocx.openxml.wordprocessing.paragraph_properties import ParagraphProperties  # noqa
from pydocx.openxml.wordprocessing.run import Run
from pydocx.openxml.wordprocessing.sdt_run import SdtRun
from pydocx.openxml.wordprocessing.simple_field import SimpleField
from pydocx.openxml.wordprocessing.smart_tag_run import SmartTagRun
from pydocx.openxml.wordprocessing.tab_char import TabChar
from pydocx.openxml.wordprocessing.text import Text
from pydocx.openxml.wordprocessing.table_cell import TableCell
from pydocx.util.memoize import memoized


class Paragraph(XmlModel):
    XML_TAG = 'p'

    properties = XmlChild(type=ParagraphProperties)

    children = XmlCollection(
        Run,
        Hyperlink,
        SmartTagRun,
        InsertedRun,
        DeletedRun,
        SdtRun,
        SimpleField,
        Bookmark
    )

    def __init__(self, **kwargs):
        super(Paragraph, self).__init__(**kwargs)
        self._effective_properties = None

    @property
    def is_empty(self):
        if not self.children:
            return True

        # we may have cases when a paragraph has a Bookmark with name '_GoBack'
        # and we should treat it as empty paragraph
        if len(self.children) == 1:
            first_child = self.children[0]
            if isinstance(first_child, Bookmark) and \
                            first_child.name in ('_GoBack',):
                return True
            # We can have cases when only run properties are defined and no text
            elif not first_child.children:
                return True
        return False

    @property
    def effective_properties(self):
        # TODO need to calculate effective properties like Run
        if not self._effective_properties:
            properties = self.properties
            self._effective_properties = properties
        return self._effective_properties

    @property
    def numbering_definition(self):
        return self.get_numbering_definition()

    def has_structured_document_parent(self):
        from pydocx.openxml.wordprocessing import SdtBlock
        return self.has_ancestor(SdtBlock)

    def get_style_chain_stack(self):
        if not self.properties:
            return

        parent_style = self.properties.parent_style
        if not parent_style:
            return

        # TODO the getattr is necessary because of footnotes. From the context
        # of a footnote, a paragraph's container is the footnote part, which
        # doesn't have access to the style_definitions_part
        part = getattr(self.container, 'style_definitions_part', None)
        if part:
            style_stack = part.get_style_chain_stack('paragraph', parent_style)
            for result in style_stack:
                yield result

    @property
    def heading_style(self):
        if hasattr(self, '_heading_style'):
            return getattr(self, '_heading_style')
        style_stack = self.get_style_chain_stack()
        heading_style = None
        for style in style_stack:
            if style.is_a_heading():
                heading_style = style
                break
        self.heading_style = heading_style
        return heading_style

    @heading_style.setter
    def heading_style(self, style):
        self._heading_style = style

    @memoized
    def get_numbering_definition(self):
        # TODO the getattr is necessary because of footnotes. From the context
        # of a footnote, a paragraph's container is the footnote part, which
        # doesn't have access to the numbering_definitions_part
        part = getattr(self.container, 'numbering_definitions_part', None)
        if not part:
            return
        if not self.effective_properties:
            return
        numbering_properties = self.effective_properties.numbering_properties
        if not numbering_properties:
            return
        return part.numbering.get_numbering_definition(
            num_id=numbering_properties.num_id,
        )

    @memoized
    def get_numbering_level(self):
        numbering_definition = self.get_numbering_definition()
        if not numbering_definition:
            return
        if not self.effective_properties:
            return
        numbering_properties = self.effective_properties.numbering_properties
        if not numbering_properties:
            return
        return numbering_definition.get_level(
            level_id=numbering_properties.level_id,
        )

    @property
    def runs(self):
        for p_child in self.children:
            if isinstance(p_child, Run):
                yield p_child

    @property
    def bookmark_name(self):
        for p_child in self.children:
            if isinstance(p_child, Bookmark):
                return p_child.name

    def get_text(self, tab_char=None):
        '''
        Return a string of all of the contained Text nodes concatenated
        together. If `tab_char` is set, then any TabChar encountered will be
        represented in the returned text using the specified string.

        For example:

        Given the following paragraph XML definition:

            <p>
                <r>
                    <t>abc</t>
                </r>
                <r>
                    <t>def</t>
                </r>
            </p>

        `get_text()` will return 'abcdef'
        '''

        text = []
        for run in self.runs:
            for r_child in run.children:
                if isinstance(r_child, Text):
                    if r_child.text:
                        text.append(r_child.text)
                if tab_char and isinstance(r_child, TabChar):
                    text.append(tab_char)
        return ''.join(text)

    def get_number_of_initial_tabs(self):
        '''
        Return the number of initial TabChars.
        '''
        tab_count = 0
        for p_child in self.children:
            if isinstance(p_child, Run):
                for r_child in p_child.children:
                    if isinstance(r_child, TabChar):
                        tab_count += 1
                    else:
                        break
            else:
                break
        return tab_count

    @property
    @memoized
    def has_numbering_properties(self):
        return bool(getattr(self.effective_properties, 'numbering_properties', None))

    @property
    @memoized
    def has_numbering_definition(self):
        return bool(self.numbering_definition)

    @property
    @memoized
    def has_border_properties(self):
        return bool(getattr(self.effective_properties, 'border_properties', None))

    def get_indentation(self, indentation, only_level_ind=False):
        '''
        Get specific indentation of the current paragraph. If indentation is
        not present on the paragraph level, get it from the numbering definition.
        '''

        ind = None

        if self.properties:
            if not only_level_ind:
                ind = self.properties.to_int(indentation)
            if ind is None:
                level = self.get_numbering_level()
                ind = level.paragraph_properties.to_int(indentation, default=0)

        return ind

    def have_same_numbering_properties_as(self, paragraph):
        prop1 = getattr(self.effective_properties, 'numbering_properties', None)
        prop2 = getattr(paragraph.effective_properties, 'numbering_properties', None)

        if prop1 == prop2:
            return True

        return False

    def get_spacing(self):
        """Get paragraph spacing according to:
                ECMA-376, 3rd Edition (June, 2011),
                Fundamentals and Markup Language Reference § 17.3.1.33.

            Note: Partial implementation for now.
        """
        results = {
            'line': None,
            'after': None,
            'before': None,
            'contextual_spacing': False,
            'parent_style': None
        }

        # Get the paragraph_properties from the parent styles
        style_paragraph_properties = None
        for style in self.get_style_chain_stack():
            if style.paragraph_properties:
                style_paragraph_properties = style.paragraph_properties
                break

        if style_paragraph_properties:
            results['contextual_spacing'] = bool(style_paragraph_properties.contextual_spacing)

        default_paragraph_properties = None
        if self.default_doc_styles and self.default_doc_styles.paragraph:
            default_paragraph_properties = self.default_doc_styles.paragraph.properties

        # Spacing properties can be defined in multiple places and we need to get some
        # kind of order of check
        properties_order = [None, None, None]
        if self.properties:
            properties_order[0] = self.properties
        if isinstance(self.parent, TableCell):
            properties_order[1] = self.parent.parent_table.get_paragraph_properties()
        if not self.properties or not self.properties.spacing_properties:
            properties_order[2] = default_paragraph_properties

        spacing_properties = None
        contextual_spacing = None

        for properties in properties_order:
            if spacing_properties is None:
                spacing_properties = getattr(properties, 'spacing_properties', None)
            if contextual_spacing is None:
                contextual_spacing = getattr(properties, 'contextual_spacing', None)

        if not spacing_properties:
            return results

        if contextual_spacing is not None:
            results['contextual_spacing'] = bool(contextual_spacing)

        if self.properties:
            results['parent_style'] = self.properties.parent_style

        spacing_line = spacing_properties.to_int('line')
        spacing_after = spacing_properties.to_int('after')
        spacing_before = spacing_properties.to_int('before')

        if default_paragraph_properties and spacing_line is None \
                and bool(spacing_properties.after_auto_spacing):
            # get the spacing_line from the default definition
            spacing_line = default_paragraph_properties.spacing_properties.to_int('line')

        if spacing_line:
            line = spacing_line / 240.0
            # default line spacing is 1 so no need to add attribute
            if line != 1.0:
                results['line'] = line

        if spacing_after is not None:
            results['after'] = spacing_after

        if spacing_before is not None:
            results['before'] = spacing_before

        return results
