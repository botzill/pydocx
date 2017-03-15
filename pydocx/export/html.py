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
    EMUS_PER_PIXEL,
    HTML_WHITE_SPACE
)
from pydocx.export.base import PyDocXExporter
from pydocx.export.numbering_span import NumberingItem
from pydocx.openxml import wordprocessing
from pydocx.util.uri import uri_is_external
from pydocx.util.xml import convert_dictionary_to_style_fragment
from pydocx.export.html_tag import (
    HtmlTag,
    is_only_whitespace
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
            },
            'p': {
                'margin-top': '0',
                'margin-bottom': '0'
            },
            'ol': {
                'margin-top': '0',
                'margin-bottom': '0'
            },
            'ul': {
                'margin-top': '0',
                'margin-bottom': '0'
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
        if isinstance(paragraph.parent, wordprocessing.TableCell):
            cell_properties = paragraph.parent.properties
            if cell_properties and cell_properties.is_continue_vertical_merge:
                # We ignore such paragraphs here because are added via rowspan
                return
        if paragraph.is_empty:
            return HtmlTag('p', custom_text=HTML_WHITE_SPACE)

        heading_style = paragraph.heading_style
        if heading_style:
            tag = self.get_heading_tag(paragraph)
            if tag:
                return tag

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

        for result in self.border_and_shading_builder.export_borders(
                run, results, first_pass=self.first_pass):
            yield result

    def export_paragraph(self, paragraph):
        results = super(PyDocXHTMLExporter, self).export_paragraph(paragraph)

        tag = self.get_paragraph_tag(paragraph)
        if tag:
            attrs = self.get_paragraph_styles(paragraph)
            tag.attrs.update(attrs)
            results = tag.apply(results)

        for result in self.border_and_shading_builder.export_borders(
                paragraph, results, first_pass=self.first_pass):
            yield result

    def export_paragraph_property_justification(self, paragraph, results):
        # TODO these classes could be applied on the paragraph, and not as
        # inline spans
        # TODO These alignment values are for traditional conformance. Strict
        # conformance uses different values
        attrs = self.get_paragraph_property_justification(paragraph)
        if attrs:
            tag = HtmlTag('span', **attrs)
            results = tag.apply(results, allow_empty=False)
        return results

    def get_paragraph_property_justification(self, paragraph):
        attrs = {}
        if not paragraph.effective_properties:
            return attrs

        alignment = paragraph.effective_properties.justification

        if alignment in [JUSTIFY_LEFT, JUSTIFY_CENTER, JUSTIFY_RIGHT]:
            pydocx_class = 'pydocx-{alignment}'.format(
                alignment=alignment,
            )
            attrs = {
                'class': pydocx_class,
            }
        elif alignment is not None:
            # TODO What if alignment is something else?
            pass

        return attrs

    def export_paragraph_property_indentation(self, paragraph, results):
        # TODO these classes should be applied on the paragraph, and not as
        # inline styles

        attrs = self.get_paragraph_property_indentation(paragraph)

        if attrs:
            tag = HtmlTag('span', **attrs)
            results = tag.apply(results, allow_empty=False)

        return results

    def get_paragraph_property_spacing(self, paragraph):
        style = {}
        if self.first_pass:
            return style

        try:
            current_par_index = self.paragraphs.index(paragraph)
        except ValueError:
            return style

        previous_paragraph = None
        next_paragraph = None
        previous_paragraph_spacing = None
        next_paragraph_spacing = None
        spacing_after = None
        spacing_before = None

        current_paragraph_spacing = paragraph.get_spacing()

        if current_par_index > 0:
            previous_paragraph = self.paragraphs[current_par_index - 1]
            previous_paragraph_spacing = previous_paragraph.get_spacing()
        if current_par_index < len(self.paragraphs) - 1:
            next_paragraph = self.paragraphs[current_par_index + 1]
            next_paragraph_spacing = next_paragraph.get_spacing()

        if next_paragraph:
            current_after = current_paragraph_spacing['after'] or 0
            next_before = next_paragraph_spacing['before'] or 0

            same_style = current_paragraph_spacing['parent_style'] == \
                next_paragraph_spacing['parent_style']

            if same_style:
                if not current_paragraph_spacing['contextual_spacing']:
                    if next_paragraph_spacing['contextual_spacing']:
                        spacing_after = current_after
                    else:
                        if current_after > next_before:
                            spacing_after = current_after
            else:
                if current_after > next_before:
                    spacing_after = current_after
        else:
            spacing_after = current_paragraph_spacing['after']

        if previous_paragraph:
            current_before = current_paragraph_spacing['before'] or 0
            prev_after = previous_paragraph_spacing['after'] or 0

            same_style = current_paragraph_spacing['parent_style'] == \
                previous_paragraph_spacing['parent_style']

            if same_style:
                if not current_paragraph_spacing['contextual_spacing']:
                    if previous_paragraph_spacing['contextual_spacing']:
                        if current_before > prev_after:
                            spacing_before = current_before - prev_after
                        else:
                            spacing_before = 0
                    else:
                        if current_before > prev_after:
                            spacing_before = current_before
            else:
                if current_before > prev_after:
                    spacing_before = current_before
        else:
            spacing_before = current_paragraph_spacing['before']

        if current_paragraph_spacing['line']:
            style['line-height'] = '{0}%'.format(current_paragraph_spacing['line'] * 100)

        if spacing_after:
            style['margin-bottom'] = '{0:.2f}em'.format(convert_twips_to_ems(spacing_after))

        if spacing_before:
            style['margin-top'] = '{0:.2f}em'.format(convert_twips_to_ems(spacing_before))

        if style:
            style = {
                'style': convert_dictionary_to_style_fragment(style)
            }

        return style

    def get_paragraph_property_indentation(self, paragraph):
        style = {}
        attrs = {}
        properties = paragraph.effective_properties

        indentation_right = None
        indentation_left = 0
        indentation_first_line = None
        span_paragraph_properties = None
        span_indentation_left = 0

        try:
            if isinstance(paragraph.parent, NumberingItem):
                span_paragraph_properties = paragraph.parent.numbering_span.numbering_level.\
                    paragraph_properties
                span_indentation_left = span_paragraph_properties.to_int(
                    'indentation_left',
                    default=0
                )
                span_indentation_hanging = span_paragraph_properties.to_int(
                    'indentation_hanging',
                    default=0
                )
                if span_paragraph_properties:
                    indentation_left -= (span_indentation_left - span_indentation_hanging)

        except AttributeError:
            pass

        if properties:
            indentation_right = properties.to_int('indentation_right')

            if properties.numbering_properties is None:
                # For paragraph inside list we need to properly adjust indentations
                # by recalculating their indentations based on the parent span
                indentation_left = properties.to_int('indentation_left', default=0)
                indentation_first_line = properties.to_int('indentation_first_line', default=0)

                if isinstance(paragraph.parent, NumberingItem):
                    if properties.is_list_paragraph and properties.no_indentation:
                        indentation_left = 0
                    elif span_paragraph_properties:
                        indentation_left -= span_indentation_left
                        # In this case we don't need to set text-indent separately because
                        # it's part of the left margin
                        indentation_left += indentation_first_line
                        indentation_first_line = None
                    else:
                        # TODO Here we may encounter fake lists and not always margins are
                        # set properly.
                        pass
            else:
                indentation_left = None
                indentation_first_line = None
                paragraph_num_level = paragraph.get_numbering_level()

                if paragraph_num_level:
                    listing_style = self.export_listing_paragraph_property_indentation(
                        paragraph,
                        paragraph_num_level.paragraph_properties,
                        include_text_indent=True
                    )
                    if 'text-indent' in listing_style and \
                            listing_style['text-indent'] != '0.00em':
                        style['text-indent'] = listing_style['text-indent']

        if indentation_right:
            right = convert_twips_to_ems(indentation_right)
            style['margin-right'] = '{0:.2f}em'.format(right)

        if indentation_left:
            left = convert_twips_to_ems(indentation_left)
            style['margin-left'] = '{0:.2f}em'.format(left)

        if indentation_first_line:
            first_line = convert_twips_to_ems(indentation_first_line)
            style['text-indent'] = '{0:.2f}em'.format(first_line)

        if style:
            attrs = {
                'style': convert_dictionary_to_style_fragment(style)
            }

        return attrs

    def get_paragraph_styles(self, paragraph):
        attributes = {}

        property_rules = [
            (True, self.get_paragraph_property_justification),
            (True, self.get_paragraph_property_indentation),
            (True, self.get_paragraph_property_spacing),
        ]
        for actual_value, handler in property_rules:
            if actual_value:
                handler_results = handler(paragraph)
                for attr_name in ['style', 'class']:
                    new_value = handler_results.get(attr_name, '')
                    if new_value:
                        if attr_name in attributes:
                            attributes[attr_name] += ';%s' % new_value
                        else:
                            attributes[attr_name] = '%s' % new_value

        return attributes

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

        # We don't allow negative hanging
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
        for result in self.border_and_shading_builder.export_close_paragraph_border():
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

        results = self.yield_nested(
            numbering_spans,
            self.export_node
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
        results = super(PyDocXHTMLExporter, self).export_numbering_item(numbering_item)

        style = None

        if numbering_item.children:
            level_properties = numbering_item.numbering_span. \
                numbering_level.paragraph_properties
            # get the first paragraph properties which will contain information
            # on how to properly indent listing item
            paragraph = numbering_item.get_first_child()

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
