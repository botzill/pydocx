# coding: utf-8
from __future__ import (
    absolute_import,
    division,
    print_function,
    unicode_literals,
)

import base64
import posixpath
from itertools import chain

from pydocx.constants import (
    JUSTIFY_CENTER,
    JUSTIFY_LEFT,
    JUSTIFY_RIGHT,
    POINTS_PER_EM,
    PYDOCX_STYLES,
    TWIPS_PER_POINT,
    EMUS_PER_PIXEL
)
from pydocx.export.base import PyDocXExporter
from pydocx.export.numbering_span import NumberingItem
from pydocx.openxml import wordprocessing
from pydocx.util.uri import uri_is_external
from pydocx.util.xml import (
    convert_dictionary_to_html_attributes,
    convert_dictionary_to_style_fragment,
)


def convert_twips_to_ems(value):
    '''
    >>> convert_twips_to_ems(30)
    0.125
    '''
    return value / TWIPS_PER_POINT / POINTS_PER_EM


def convert_emus_to_pixels(emus):
    return emus / EMUS_PER_PIXEL


def get_first_from_sequence(sequence, default=None):
    '''
    Given a sequence, return the first item in the sequence. If the sequence is
    empty, return the passed in default.
    '''
    first_result = default
    try:
        first_result = next(sequence)
    except StopIteration:
        pass
    return first_result


def is_only_whitespace(obj):
    '''
    If the obj has `strip` return True if calling strip on the obj results in
    an empty instance. Otherwise, return False.
    '''
    if hasattr(obj, 'strip'):
        return not obj.strip()
    return False


def is_not_empty_and_not_only_whitespace(gen):
    '''
    Determine if a generator is empty, or consists only of whitespace.

    If the generator is non-empty, return the original generator. Otherwise,
    return None
    '''
    queue = []
    if gen is None:
        return
    try:
        for item in gen:
            queue.append(item)
            is_whitespace = True
            if isinstance(item, HtmlTag):
                # If we encounter a tag that allows whitespace, then we can stop
                is_whitespace = not item.allow_whitespace
            else:
                is_whitespace = is_only_whitespace(item)

            if not is_whitespace:
                # This item isn't whitespace, so we're done scanning
                return chain(queue, gen)

    except StopIteration:
        pass


class HtmlTag(object):
    closed_tag_format = '</{tag}>'

    def __init__(
            self,
            tag,
            allow_self_closing=False,
            closed=False,
            allow_whitespace=False,
            **attrs
    ):
        self.tag = tag
        self.allow_self_closing = allow_self_closing
        self.attrs = attrs
        self.closed = closed
        self.allow_whitespace = allow_whitespace

    def apply(self, results, allow_empty=True):
        if not allow_empty:
            results = is_not_empty_and_not_only_whitespace(results)
            if results is None:
                return

        sequence = [[self]]
        if results is not None:
            sequence.append(results)

        if not self.allow_self_closing:
            sequence.append([self.close()])

        results = chain(*sequence)

        for result in results:
            yield result

    def close(self):
        return HtmlTag(
            tag=self.tag,
            closed=True,
        )

    def to_html(self):
        if self.closed is True:
            return self.closed_tag_format.format(tag=self.tag)
        else:
            attrs = self.get_html_attrs()
            end_bracket = '>'
            if self.allow_self_closing:
                end_bracket = ' />'
            if attrs:
                return '<{tag} {attrs}{end}'.format(
                    tag=self.tag,
                    attrs=attrs,
                    end=end_bracket,
                )
            else:
                return '<{tag}{end}'.format(tag=self.tag, end=end_bracket)

    def get_html_attrs(self):
        return convert_dictionary_to_html_attributes(self.attrs)


class PyDocXHTMLExporter(PyDocXExporter):
    def __init__(self, *args, **kwargs):
        super(PyDocXHTMLExporter, self).__init__(*args, **kwargs)
        self.table_cell_rowspan_tracking = {}
        self.in_table_cell = False
        self.heading_level_conversion_map = {
            'heading 1': 'h1',
            'heading 2': 'h2',
            'heading 3': 'h3',
            'heading 4': 'h4',
            'heading 5': 'h5',
            'heading 6': 'h6',
        }
        self.default_heading_level = 'h6'

    def head(self):
        tag = HtmlTag('head')
        results = chain(self.meta(), self.style())
        return tag.apply(results)

    def style(self):
        styles = {
            'body': {
                'margin': '0px auto',
            }
        }

        if self.page_width:
            width = self.page_width / POINTS_PER_EM
            styles['body']['width'] = '%.2fem' % width

        result = []
        for name, definition in sorted(PYDOCX_STYLES.items()):
            result.append('.pydocx-%s {%s}' % (
                name,
                convert_dictionary_to_style_fragment(definition),
            ))

        for name, definition in sorted(styles.items()):
            result.append('%s {%s}' % (
                name,
                convert_dictionary_to_style_fragment(definition),
            ))

        tag = HtmlTag('style')
        return tag.apply(''.join(result))

    def meta(self):
        yield HtmlTag('meta', charset='utf-8', allow_self_closing=True)

    def export(self):
        return ''.join(
            result.to_html() if isinstance(result, HtmlTag)
            else result
            for result in super(PyDocXHTMLExporter, self).export()
        )

    def export_document(self, document):
        tag = HtmlTag('html')
        results = super(PyDocXHTMLExporter, self).export_document(document)
        sequence = []
        head = self.head()
        if head is not None:
            sequence.append(head)
        if results is not None:
            sequence.append(results)
        return tag.apply(chain(*sequence))

    def export_body(self, body):
        results = super(PyDocXHTMLExporter, self).export_body(body)
        tag = HtmlTag('body')
        return tag.apply(chain(results, self.footer()))

    def footer(self):
        for result in self.export_footnotes():
            yield result

    def export_footnotes(self):
        results = super(PyDocXHTMLExporter, self).export_footnotes()
        attrs = {
            'class': 'pydocx-list-style-type-decimal',
        }
        ol = HtmlTag('ol', **attrs)
        results = ol.apply(results, allow_empty=False)

        page_break = HtmlTag('hr', allow_self_closing=True)
        return page_break.apply(results, allow_empty=False)

    def export_footnote(self, footnote):
        results = super(PyDocXHTMLExporter, self).export_footnote(footnote)
        tag = HtmlTag('li')
        return tag.apply(results, allow_empty=False)

    def get_paragraph_tag(self, paragraph):
        heading_style = paragraph.heading_style
        if heading_style:
            tag = self.get_heading_tag(paragraph)
            if tag:
                return tag
        if self.in_table_cell:
            return
        if paragraph.has_structured_document_parent():
            return
        if isinstance(paragraph.parent, NumberingItem):
            return
        return HtmlTag('p')

    def get_heading_tag(self, paragraph):
        if paragraph.has_ancestor(NumberingItem):
            # Force-bold headings that appear in list items
            return HtmlTag('strong')
        heading_style = paragraph.heading_style
        tag = self.heading_level_conversion_map.get(
            heading_style.name.lower(),
            self.default_heading_level,
        )
        if paragraph.bookmark_name:
            return HtmlTag(tag, id=paragraph.bookmark_name)
        return HtmlTag(tag)

    def export_run(self, run):
        results = super(PyDocXHTMLExporter, self).export_run(run)

        for result in self.export_borders(run, results, tag_name='span'):
            yield result

    def export_paragraph(self, paragraph):
        results = super(PyDocXHTMLExporter, self).export_paragraph(paragraph)
        results = is_not_empty_and_not_only_whitespace(results)

        # TODO@botzill In PR#234 we render empty paragraphs properly so
        # we don't need this check anymore. Adding for now and to be removed when merging
        if results is None and not paragraph.has_border_properties:
            return

        tag = self.get_paragraph_tag(paragraph)
        if tag:
            results = tag.apply(results)

        for tag in self.export_borders(paragraph, results, tag_name='div'):
            yield tag

    def export_close_paragraph_border(self):
        if self.current_border_item.get('Paragraph'):
            yield HtmlTag('div', closed=True)
            self.current_border_item['Paragraph'] = None

    def export_borders(self, item, results, tag_name='div'):
        if self.first_pass:
            for result in results:
                yield result
            return

        # For now we have here Paragraph and Run
        item_name = item.__class__.__name__
        item_is_run = isinstance(item, wordprocessing.Run)
        item_is_paragraph = isinstance(item, wordprocessing.Paragraph)

        prev_borders_properties = None
        prev_shading_properties = None

        border_properties = None
        shading_properties = None

        current_border_item = self.current_border_item.get(item_name)
        if current_border_item:
            item_properties = current_border_item.effective_properties
            prev_borders_properties = item_properties.border_properties
            prev_shading_properties = item_properties.shading_properties

        last_item = False
        close_border = True

        def prev_properties():
            return prev_borders_properties or prev_shading_properties

        def current_properties():
            return border_properties or shading_properties

        def current_item_is_last_child(children, child_type):
            for p_child in reversed(children):
                if isinstance(p_child, child_type):
                    return p_child == item
            return False

        def is_last_item():
            if item_is_paragraph:
                if isinstance(item.parent, wordprocessing.TableCell):
                    return current_item_is_last_child(
                        item.parent.children, wordprocessing.Paragraph)
                elif item == self.last_paragraph:
                    return True
            elif item_is_run:
                # Check if current item is the last Run item from paragraph children
                return current_item_is_last_child(item.parent.children, wordprocessing.Run)

            return False

        if item.effective_properties:
            border_properties = item.effective_properties.border_properties
            shading_properties = item.effective_properties.shading_properties

            if current_properties():
                last_item = is_last_item()
                close_border = False
                run_has_different_parent = False

                # If run is from different paragraph then we may need to draw separate border
                # even if border properties are the same
                if item_is_run and current_border_item:
                    if current_border_item.parent != item.parent:
                        run_has_different_parent = True

                if border_properties != prev_borders_properties or \
                    shading_properties != prev_shading_properties or \
                        run_has_different_parent:
                    if prev_properties() is not None:
                        # We have a previous border tag opened, so need to close it
                        yield HtmlTag(tag_name, closed=True)

                    # Open a new tag for the new border/shading and include all the properties
                    attrs = self.get_borders_property(
                        border_properties,
                        prev_borders_properties,
                        shading_properties
                    )
                    yield HtmlTag(tag_name, closed=False, **attrs)

                    self.current_border_item[item_name] = item

                if border_properties == prev_borders_properties:
                    border_between = getattr(border_properties, 'between', None)
                    add_between_border = bool(border_between)

                    if border_between and prev_borders_properties is not None:
                        if shading_properties:
                            if shading_properties == prev_shading_properties:
                                add_between_border = True
                            else:

                                add_between_border = prev_borders_properties.bottom != \
                                                     border_between

                    if add_between_border:
                        # Render border between items
                        border_attrs = self.get_borders_property(
                            border_properties,
                            prev_borders_properties,
                            shading_properties,
                            only_between=True)

                        yield HtmlTag(tag_name, **border_attrs)
                        yield HtmlTag(tag_name, closed=True)

        if close_border and prev_properties() is not None:
            # At this stage we need to make sure that if there is an previously open tag
            # about border we need to close it
            yield HtmlTag(tag_name, closed=True)
            self.current_border_item[item_name] = None

        # All the inner items inside border tag are issued here
        for result in results:
            yield result

        if current_properties() and last_item:
            # If the item with border is the last one we need to make sure that we close the
            # tag
            yield HtmlTag(tag_name, closed=True)
            self.current_border_item[item_name] = None

    def get_borders_property(
            self,
            border_properties,
            prev_border_properties,
            shading_properties=None,
            only_between=False):
        attrs = {}
        style = {}

        if border_properties:
            if only_between:
                style.update(border_properties.get_between_border_style())
            else:
                style.update(border_properties.get_padding_style())
                style.update(border_properties.get_shadow_style())

                border_style = border_properties.get_border_style()

                if prev_border_properties and \
                        isinstance(prev_border_properties, wordprocessing.ParagraphBorders):

                    cur_top = border_properties.top
                    prev_bottom = prev_border_properties.bottom

                    all_borders_defined = all([
                        border_properties.borders_have_same_properties(),
                        prev_border_properties.borders_have_same_properties()
                    ])
                    # We need to reset one border if adjacent identical borders are met
                    if all_borders_defined and cur_top == prev_bottom:
                        border_style['border-top'] = '0'
                style.update(border_style)

        if shading_properties and shading_properties.background_color:
            style['background-color'] = '#{0}'.format(shading_properties.background_color)

        if style:
            attrs['style'] = convert_dictionary_to_style_fragment(style)

        return attrs

    def export_paragraph_property_justification(self, paragraph, results):
        # TODO these classes could be applied on the paragraph, and not as
        # inline spans
        alignment = paragraph.effective_properties.justification
        # TODO These alignment values are for traditional conformance. Strict
        # conformance uses different values
        if alignment in [JUSTIFY_LEFT, JUSTIFY_CENTER, JUSTIFY_RIGHT]:
            pydocx_class = 'pydocx-{alignment}'.format(
                alignment=alignment,
            )
            attrs = {
                'class': pydocx_class,
            }
            tag = HtmlTag('span', **attrs)
            results = tag.apply(results, allow_empty=False)
        elif alignment is not None:
            # TODO What if alignment is something else?
            pass
        return results

    def export_paragraph_property_indentation(self, paragraph, results):
        # TODO these classes should be applied on the paragraph, and not as
        # inline styles

        properties = paragraph.effective_properties

        style = {}

        # Numbering properties can define a text indentation on a paragraph
        if properties.numbering_properties:
            indentation_left = None
            indentation_first_line = None

            paragraph_num_level = paragraph.get_numbering_level()

            if paragraph_num_level:
                listing_style = self.export_listing_paragraph_property_indentation(
                    paragraph,
                    paragraph_num_level.paragraph_properties,
                    include_text_indent=True
                )
                if 'text-indent' in listing_style and listing_style['text-indent'] != '0.00em':
                    style['text-indent'] = listing_style['text-indent']
                    style['display'] = 'inline-block'
        else:
            indentation_left = properties.to_int('indentation_left')
            indentation_first_line = properties.to_int('indentation_first_line')

        indentation_right = properties.to_int('indentation_right')

        if indentation_right:
            right = convert_twips_to_ems(indentation_right)
            style['margin-right'] = '{0:.2f}em'.format(right)

        if indentation_left:
            left = convert_twips_to_ems(indentation_left)
            style['margin-left'] = '{0:.2f}em'.format(left)

        if indentation_first_line:
            first_line = convert_twips_to_ems(indentation_first_line)
            style['text-indent'] = '{0:.2f}em'.format(first_line)
            style['display'] = 'inline-block'

        if style:
            attrs = {
                'style': convert_dictionary_to_style_fragment(style)
            }
            tag = HtmlTag('span', **attrs)
            results = tag.apply(results, allow_empty=False)

        return results

    def export_listing_paragraph_property_indentation(
            self,
            paragraph,
            level_properties,
            include_text_indent=False
    ):
        style = {}

        if not level_properties or not paragraph.has_numbering_properties:
            return style

        level_indentation_step = \
            paragraph.numbering_definition.get_indentation_between_levels()

        paragraph_properties = paragraph.properties

        level_ind_left = level_properties.to_int('indentation_left', default=0)
        level_ind_hanging = level_properties.to_int('indentation_hanging', default=0)

        paragraph_ind_left = paragraph_properties.to_int('indentation_left', default=0)
        paragraph_ind_hanging = paragraph_properties.to_int('indentation_hanging', default=0)
        paragraph_ind_first_line = paragraph_properties.to_int('indentation_first_line',
                                                               default=0)

        left = paragraph_ind_left or level_ind_left
        hanging = paragraph_ind_hanging or level_ind_hanging
        # At this point we have no info about indentation, so we keep the default one
        if not left and not hanging:
            return style

        # All the bellow left margin calculation is done because html ul/ol/li elements have
        # their default indentations and we need to make sure that we migrate as near as
        # possible solution to html.
        margin_left = left

        # Because hanging can be set independently, we remove it from left margin and will
        # be added as text-indent later on
        margin_left -= hanging

        # Take into account that current span can have custom left margin
        if level_indentation_step > level_ind_hanging:
            margin_left -= (level_indentation_step - level_ind_hanging)
        else:
            margin_left -= level_indentation_step

        # First line are added to left margins
        margin_left += paragraph_ind_first_line

        if isinstance(paragraph.parent, NumberingItem):
            try:
                # In case of nested lists elements, we need to adjust left margin
                # based on the parent item
                parent_paragraph = paragraph.parent.numbering_span.parent.get_first_child()

                parent_ind_left = parent_paragraph.get_indentation('indentation_left')
                parent_ind_hanging = parent_paragraph.get_indentation('indentation_hanging')
                parent_lvl_ind_hanging = parent_paragraph.get_indentation(
                    'indentation_hanging')

                margin_left -= (parent_ind_left - parent_ind_hanging)
                margin_left -= parent_lvl_ind_hanging
                # To mimic the word way of setting first line, we need to move back(left) all
                # elements by first_line value
                margin_left -= parent_paragraph.get_indentation('indentation_first_line')
            except AttributeError:
                pass

        # Here as well, we remove the default hanging which word adds
        # because <li> tag will provide it's own
        hanging -= level_ind_hanging

        if margin_left:
            margin_left = convert_twips_to_ems(margin_left)
            style['margin-left'] = '{0:.2f}em'.format(margin_left)

        # we don't allow negative hanging
        if hanging < 0:
            hanging = 0

        if include_text_indent:
            if hanging is not None:
                # Now, here we add the hanging as text-indent for the paragraph
                hanging = convert_twips_to_ems(hanging)
                style['text-indent'] = '{0:.2f}em'.format(hanging)

        return style

    def get_run_styles_to_apply(self, run):
        parent_paragraph = run.get_first_ancestor(wordprocessing.Paragraph)
        if parent_paragraph and parent_paragraph.heading_style:
            results = self.get_run_styles_to_apply_for_heading(run)
        else:
            results = super(PyDocXHTMLExporter, self).get_run_styles_to_apply(run)
        for result in results:
            yield result

    def get_run_styles_to_apply_for_heading(self, run):
        allowed_handlers = set([
            self.export_run_property_italic,
            self.export_run_property_hidden,
            self.export_run_property_vanish,
        ])

        handlers = super(PyDocXHTMLExporter, self).get_run_styles_to_apply(run)
        for handler in handlers:
            if handler in allowed_handlers:
                yield handler

    def export_run_property(self, tag, run, results):
        # Any leading whitespace in the run is not styled.
        for result in results:
            if is_only_whitespace(result):
                yield result
            else:
                # We've encountered something that isn't explicit whitespace
                results = chain([result], results)
                break
        else:
            results = None

        if results:
            for result in tag.apply(results):
                yield result

    def export_run_property_bold(self, run, results):
        tag = HtmlTag('strong')
        return self.export_run_property(tag, run, results)

    def export_run_property_italic(self, run, results):
        tag = HtmlTag('em')
        return self.export_run_property(tag, run, results)

    def export_run_property_underline(self, run, results):
        attrs = {
            'class': 'pydocx-underline',
        }
        tag = HtmlTag('span', **attrs)
        return self.export_run_property(tag, run, results)

    def export_run_property_caps(self, run, results):
        attrs = {
            'class': 'pydocx-caps',
        }
        tag = HtmlTag('span', **attrs)
        return self.export_run_property(tag, run, results)

    def export_run_property_small_caps(self, run, results):
        attrs = {
            'class': 'pydocx-small-caps',
        }
        tag = HtmlTag('span', **attrs)
        return self.export_run_property(tag, run, results)

    def export_run_property_dstrike(self, run, results):
        attrs = {
            'class': 'pydocx-strike',
        }
        tag = HtmlTag('span', **attrs)
        return self.export_run_property(tag, run, results)

    def export_run_property_strike(self, run, results):
        attrs = {
            'class': 'pydocx-strike',
        }
        tag = HtmlTag('span', **attrs)
        return self.export_run_property(tag, run, results)

    def export_run_property_vanish(self, run, results):
        attrs = {
            'class': 'pydocx-hidden',
        }
        tag = HtmlTag('span', **attrs)
        return self.export_run_property(tag, run, results)

    def export_run_property_hidden(self, run, results):
        attrs = {
            'class': 'pydocx-hidden',
        }
        tag = HtmlTag('span', **attrs)
        return self.export_run_property(tag, run, results)

    def export_run_property_vertical_align(self, run, results):
        if run.effective_properties.is_superscript():
            return self.export_run_property_vertical_align_superscript(
                run,
                results,
            )
        elif run.effective_properties.is_subscript():
            return self.export_run_property_vertical_align_subscript(
                run,
                results,
            )
        return results

    def export_run_property_vertical_align_superscript(self, run, results):
        tag = HtmlTag('sup')
        return tag.apply(results, allow_empty=False)

    def export_run_property_vertical_align_subscript(self, run, results):
        tag = HtmlTag('sub')
        return tag.apply(results, allow_empty=False)

    def export_run_property_color(self, run, results):
        if run.properties is None or run.properties.color is None:
            return results

        attrs = {
            'style': 'color:#' + run.properties.color
        }
        tag = HtmlTag('span', **attrs)
        return self.export_run_property(tag, run, results)

    def export_text(self, text):
        results = super(PyDocXHTMLExporter, self).export_text(text)
        for result in results:
            if result:
                yield result

    def export_deleted_text(self, deleted_text):
        # TODO deleted_text should be ignored if it is NOT contained within a
        # deleted run
        results = self.export_text(deleted_text)
        attrs = {
            'class': 'pydocx-delete',
        }
        tag = HtmlTag('span', **attrs)
        return tag.apply(results, allow_empty=False)

    def get_hyperlink_tag(self, target_uri):
        if target_uri:
            href = self.escape(target_uri)
            return HtmlTag('a', href=href)

    def export_hyperlink(self, hyperlink):
        results = super(PyDocXHTMLExporter, self).export_hyperlink(hyperlink)
        if not hyperlink.target_uri and hyperlink.anchor:
            tag = self.get_hyperlink_tag(target_uri='#' + hyperlink.anchor)
        else:
            tag = self.get_hyperlink_tag(target_uri=hyperlink.target_uri)
        if tag:
            results = tag.apply(results, allow_empty=False)

        # Prevent underline style from applying by temporarily monkey-patching
        # the export underline function. There's got to be a better way.
        old = self.export_run_property_underline
        self.export_run_property_underline = lambda run, results: results
        for result in results:
            yield result
        self.export_run_property_underline = old

    def get_break_tag(self, br):
        if br.is_page_break():
            tag_name = 'hr'
        else:
            tag_name = 'br'
        return HtmlTag(
            tag_name,
            allow_whitespace=True,
            allow_self_closing=True,
        )

    def export_break(self, br):
        tag = self.get_break_tag(br)
        if tag:
            yield tag

    def get_table_tag(self, table):
        return HtmlTag('table', border='1')

    def export_table(self, table):
        table_cell_spans = table.calculate_table_cell_spans()
        self.table_cell_rowspan_tracking[table] = table_cell_spans
        results = super(PyDocXHTMLExporter, self).export_table(table)

        # Before starting new table new need to make sure that if there is any paragraph
        # with border opened before, we need to close it here.
        for result in self.export_close_paragraph_border():
            yield result

        tag = self.get_table_tag(table)
        results = tag.apply(results)

        for result in results:
            yield result

    def export_table_row(self, table_row):
        results = super(PyDocXHTMLExporter, self).export_table_row(table_row)
        tag = HtmlTag('tr')
        return tag.apply(results)

    def export_table_cell(self, table_cell):
        start_new_tag = False
        colspan = 1
        if table_cell.properties:
            if table_cell.properties.should_close_previous_vertical_merge():
                start_new_tag = True
            try:
                # Should be done by properties, or as a helper
                colspan = int(table_cell.properties.grid_span)
            except (TypeError, ValueError):
                colspan = 1

        else:
            # This means the properties element is missing, which means the
            # merge element is missing
            start_new_tag = True

        tag = None
        if start_new_tag:
            parent_table = table_cell.get_first_ancestor(wordprocessing.Table)
            rowspan_counts = self.table_cell_rowspan_tracking[parent_table]
            rowspan = rowspan_counts.get(table_cell, 1)
            attrs = {}
            if colspan > 1:
                attrs['colspan'] = colspan
            if rowspan > 1:
                attrs['rowspan'] = rowspan
            tag = HtmlTag('td', **attrs)

        numbering_spans = self.yield_numbering_spans(table_cell.children)
        results = self.yield_nested_with_line_breaks_between_paragraphs(
            numbering_spans,
            self.export_node,
        )
        if tag:
            results = tag.apply(results)

        self.in_table_cell = True
        for result in results:
            yield result
        self.in_table_cell = False

    def export_drawing(self, drawing):
        length, width = drawing.get_picture_extents()
        rotate = drawing.get_picture_rotate_angle()
        relationship_id = drawing.get_picture_relationship_id()
        if not relationship_id:
            return
        image = None
        try:
            image = drawing.container.get_part_by_id(
                relationship_id=relationship_id,
            )
        except KeyError:
            pass
        attrs = {}
        if length and width:
            # The "width" in openxml is actually the height
            width_px = '{px:.0f}px'.format(px=convert_emus_to_pixels(length))
            height_px = '{px:.0f}px'.format(px=convert_emus_to_pixels(width))
            attrs['width'] = width_px
            attrs['height'] = height_px
        if rotate:
            attrs['rotate'] = rotate

        tag = self.get_image_tag(image=image, **attrs)
        if tag:
            yield tag

    def get_image_source(self, image):
        if image is None:
            return
        elif uri_is_external(image.uri):
            return image.uri
        else:
            image.stream.seek(0)
            data = image.stream.read()
            _, filename = posixpath.split(image.uri)
            extension = filename.split('.')[-1].lower()
            b64_encoded_src = 'data:image/{ext};base64,{data}'.format(
                ext=extension,
                data=base64.b64encode(data).decode(),
            )
            return self.escape(b64_encoded_src)

    def get_image_tag(self, image, width=None, height=None, rotate=None):
        image_src = self.get_image_source(image)
        if image_src:
            attrs = {
                'src': image_src
            }
            if width and height:
                attrs['width'] = width
                attrs['height'] = height
            if rotate:
                attrs['style'] = 'transform: rotate(%sdeg);' % rotate

            return HtmlTag(
                'img',
                allow_self_closing=True,
                allow_whitespace=True,
                **attrs
            )

    def export_inserted_run(self, inserted_run):
        results = super(PyDocXHTMLExporter, self).export_inserted_run(inserted_run)
        attrs = {
            'class': 'pydocx-insert',
        }
        tag = HtmlTag('span', **attrs)
        return tag.apply(results)

    def export_vml_image_data(self, image_data):
        width, height = image_data.get_picture_extents()
        if not image_data.relationship_id:
            return
        image = None
        try:
            image = image_data.container.get_part_by_id(
                relationship_id=image_data.relationship_id,
            )
        except KeyError:
            pass
        tag = self.get_image_tag(image=image, width=width, height=height)
        if tag:
            yield tag

    def export_footnote_reference(self, footnote_reference):
        results = super(PyDocXHTMLExporter, self).export_footnote_reference(
            footnote_reference,
        )
        footnote_id = footnote_reference.footnote_id
        href = '#footnote-{fid}'.format(fid=footnote_id)
        name = 'footnote-ref-{fid}'.format(fid=footnote_id)
        tag = HtmlTag('a', href=href, name=name)
        for result in tag.apply(results, allow_empty=False):
            yield result

    def export_footnote_reference_mark(self, footnote_reference_mark):
        footnote_parent = footnote_reference_mark.get_first_ancestor(
            wordprocessing.Footnote,
        )
        if not footnote_parent:
            return

        footnote_id = footnote_parent.footnote_id
        if not footnote_id:
            return

        name = 'footnote-{fid}'.format(fid=footnote_id)
        href = '#footnote-ref-{fid}'.format(fid=footnote_id)
        tag = HtmlTag('a', href=href, name=name)
        results = tag.apply(['^'])
        for result in results:
            yield result

    def export_tab_char(self, tab_char):
        results = super(PyDocXHTMLExporter, self).export_tab_char(tab_char)
        attrs = {
            'class': 'pydocx-tab',
        }
        tag = HtmlTag('span', allow_whitespace=True, **attrs)
        return tag.apply(results)

    def export_numbering_span(self, numbering_span):
        results = super(PyDocXHTMLExporter, self).export_numbering_span(numbering_span)
        pydocx_class = 'pydocx-list-style-type-{fmt}'.format(
            fmt=numbering_span.numbering_level.num_format,
        )
        attrs = {}
        tag_name = 'ul'
        if not numbering_span.numbering_level.is_bullet_format():
            attrs['class'] = pydocx_class
            tag_name = 'ol'
        tag = HtmlTag(tag_name, **attrs)
        return tag.apply(results)

    def export_numbering_item(self, numbering_item):
        results = self.yield_nested_with_line_breaks_between_paragraphs(
            numbering_item.children,
            self.export_node,
        )

        style = None

        if numbering_item.children:
            level_properties = numbering_item.numbering_span.\
                numbering_level.paragraph_properties
            # get the first paragraph properties which will contain information
            # on how to properly indent listing item
            paragraph = numbering_item.children[0]

            style = self.export_listing_paragraph_property_indentation(paragraph,
                                                                       level_properties)

        attrs = {}

        if style:
            attrs['style'] = convert_dictionary_to_style_fragment(style)

        tag = HtmlTag('li', **attrs)
        return tag.apply(results)

    def export_field_hyperlink(self, simple_field, field_args):
        results = self.yield_nested(simple_field.children, self.export_node)
        if not field_args:
            return results
        target_uri = field_args.pop(0)
        bookmark = None
        bookmark_option = False
        for arg in field_args:
            if bookmark_option is True:
                bookmark = arg
            if arg == '\l':
                bookmark_option = True
        if bookmark_option and bookmark:
            target_uri = '{0}#{1}'.format(target_uri, bookmark)

        tag = self.get_hyperlink_tag(target_uri=target_uri)
        return tag.apply(results)
