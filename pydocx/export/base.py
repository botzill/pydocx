# coding: utf-8
from __future__ import (
    absolute_import,
    division,
    print_function,
    unicode_literals,
)

import xml.sax.saxutils

from pydocx.constants import TWIPS_PER_POINT
from pydocx.exceptions import MalformedDocxException
from pydocx.export.numbering_span import (
    NumberingItem,
    NumberingSpan,
    NumberingSpanBuilder,
)
from pydocx.openxml import wordprocessing, vml, wml
from pydocx.openxml.packaging import WordprocessingDocument


class PyDocXExporter(object):
    numbering_span_builder_class = NumberingSpanBuilder

    def __init__(self, path):
        self.path = path
        self._document = None
        self._page_width = None
        self.first_pass = False

        self.footnote_tracker = []

        self.captured_runs = None
        self.complex_field_runs = []

        self.node_type_to_export_func_map = {
            wordprocessing.Document: self.export_document,
            wordprocessing.Body: self.export_body,
            wordprocessing.Paragraph: self.export_paragraph,
            wordprocessing.Run: self.export_run,
            wordprocessing.Text: self.export_text,
            wordprocessing.Hyperlink: self.export_hyperlink,
            wordprocessing.Break: self.export_break,
            wordprocessing.NoBreakHyphen: self.export_no_break_hyphen,
            wordprocessing.Table: self.export_table,
            wordprocessing.TableRow: self.export_table_row,
            wordprocessing.TableCell: self.export_table_cell,
            wordprocessing.Drawing: self.export_drawing,
            wordprocessing.SmartTagRun: self.export_smart_tag_run,
            wordprocessing.InsertedRun: self.export_inserted_run,
            wordprocessing.Picture: self.export_picture,
            wordprocessing.DeletedRun: self.export_deleted_run,
            wordprocessing.DeletedText: self.export_deleted_text,
            wordprocessing.FootnoteReference: self.export_footnote_reference,
            wordprocessing.Footnote: self.export_footnote,
            wordprocessing.FootnoteReferenceMark: self.export_footnote_reference_mark,
            wordprocessing.SdtRun: self.export_sdt_run,
            wordprocessing.SdtContentRun: self.export_sdt_content_run,
            wordprocessing.SdtBlock: self.export_sdt_block,
            wordprocessing.SdtContentBlock: self.export_sdt_content_block,
            wordprocessing.TabChar: self.export_tab_char,
            wordprocessing.FieldChar: self.export_field_char,
            wordprocessing.FieldCode: self.export_field_code,
            wordprocessing.SimpleField: self.export_simple_field,
            vml.Shape: self.export_vml_shape,
            vml.ImageData: self.export_vml_image_data,
            wml.Object: self.export_object,
            NumberingSpan: self.export_numbering_span,
            NumberingItem: self.export_numbering_item,
        }
        self.field_type_to_export_func_map = {
            'HYPERLINK': getattr(self, 'export_field_hyperlink', None),
        }

    @property
    def document(self):
        if not self._document:
            self.document = self.load_document()
        return self._document

    @document.setter
    def document(self, document):
        self._document = document

    def load_document(self):
        self.document = WordprocessingDocument(path=self.path)
        return self.document

    @property
    def main_document_part(self):
        return self.document.main_document_part

    @property
    def style_definitions_part(self):
        if self.main_document_part:
            return self.main_document_part.style_definitions_part

    @property
    def numbering_definitions_part(self):
        if self.main_document_part:
            return self.main_document_part.numbering_definitions_part

    def export(self):
        if self.main_document_part is None:
            raise MalformedDocxException
        document = self.main_document_part.document
        if document:
            # process the document in two passes, since there are some cases
            # where we can't know what to do until we look at the entire
            # document (e.g. fields)
            # In the first pass, discard any generated results
            self.first_pass = True
            for result in self.export_node(document):
                pass

            self._convert_complex_fields_into_simple_fields()

            # actually render the results
            self.first_pass = False
            for result in self.export_node(document):
                yield result

    def _convert_complex_fields_into_simple_fields(self):
        if not self.complex_field_runs:
            return

        fields = []
        field = None
        separate_triggered = False
        previous_run = None

        runs_to_remove = set()

        for run in self.complex_field_runs:
            for child in run.children:
                if field is not None and previous_run is not None:
                    if previous_run.parent is not run.parent:
                        # scope has changed
                        fields.append(field)
                        field = wordprocessing.SimpleField(
                            instr=field.instr, children=[])

                if isinstance(child, wordprocessing.FieldChar):
                    separate_triggered = False
                    runs_to_remove.add(run)
                    if child.is_type_begin():
                        field = wordprocessing.SimpleField(children=[])
                    elif child.is_type_end():
                        if field is None:
                            pass
                        else:
                            fields.append(field)
                        field = None
                    elif child.is_type_separate():
                        separate_triggered = True

                if field is None:
                    pass
                elif separate_triggered:
                    field.children.append(run)
                    runs_to_remove.add(run)
                elif isinstance(child, wordprocessing.FieldCode):
                    field.instr = child.content
                    runs_to_remove.add(run)
            previous_run = run

        for field in fields:
            if not field.children:
                continue
            first_run = field.children[0]
            previous_parent_new_children = []
            for child in first_run.parent.children:
                if child is first_run:
                    # This is the insertion point of the new field
                    previous_parent_new_children.append(field)
                elif child not in runs_to_remove:
                    previous_parent_new_children.append(child)
            first_run.parent.children = previous_parent_new_children

            for run in field.children:
                run.parent = field

    def export_node(self, node):
        caller = self.node_type_to_export_func_map.get(type(node))
        if callable(caller):
            results = caller(node)
            if results is not None:
                for result in results:
                    yield result

    @property
    def page_width(self):
        if self._page_width is None:
            page_width = self.calculate_page_width()
            if page_width is None:
                page_width = 0
            self._page_width = page_width
        return self._page_width

    def calculate_page_width(self):
        document = self.main_document_part.document
        try:
            page_size = document.body.final_section_properties.page_size
        except AttributeError:
            page_size = None
        if page_size:
            width = page_size.get('w')
            if not width:
                return
            try:
                width = int(float(width))
            except ValueError:
                return
            # pgSz is defined in twips, convert to points
            return width / TWIPS_PER_POINT

    def export_document(self, document):
        return self.export_node(document.body)

    def yield_nested(self, iterable, func):
        for item in iterable:
            for result in func(item):
                yield result

    def yield_nested_with_line_breaks_between_paragraphs(self, iterable, func):
        '''
        Yield a line break in between adjacent paragraphs, ignoring any
        paragraphs that are empty or otherwise don't contain any content.
        '''
        br = wordprocessing.Break()

        previous_was_paragraph = False
        previous_was_empty = True
        for item in iterable:
            empty = True
            is_paragraph = isinstance(item, wordprocessing.Paragraph)
            for result in func(item):
                if empty:
                    empty = False
                    if is_paragraph and previous_was_paragraph and not previous_was_empty:
                        for br_result in func(br):
                            yield br_result
                yield result
            previous_was_paragraph = is_paragraph
            if not empty:
                # Once it's not empty, there's always a previous non-empty
                # paragraph.
                previous_was_empty = empty

    def yield_numbering_spans(self, items):
        builder = self.numbering_span_builder_class(items)
        numbering_spans = builder.get_numbering_spans()
        for item in numbering_spans:
            yield item

    def export_body(self, body):
        children = self.yield_body_children(body)
        return self.yield_nested(children, self.export_node)

    def yield_body_children(self, body):
        if self.first_pass:
            return body.children
        return self.yield_numbering_spans(body.children)

    def export_paragraph(self, paragraph):
        children = self.yield_paragraph_children(paragraph)
        results = self.yield_nested(children, self.export_node)
        if paragraph.effective_properties:
            results = self.export_paragraph_apply_properties(
                paragraph, results)
        return results

    def yield_paragraph_children(self, paragraph):
        for child in paragraph.children:
            yield child

    def get_paragraph_styles_to_apply(self, paragraph):
        properties = paragraph.effective_properties
        property_rules = [
            (properties.justification, self.export_paragraph_property_justification),
            (True, self.export_paragraph_property_indentation),
        ]
        for actual_value, handler in property_rules:
            if actual_value:
                yield handler

    def export_paragraph_property_justification(self, paragraph, results):
        return results

    def export_paragraph_property_indentation(self, paragraph, results):
        return results

    def export_paragraph_apply_properties(self, paragraph, results):
        styles_to_apply = self.get_paragraph_styles_to_apply(paragraph)
        for func in styles_to_apply:
            results = func(paragraph, results)
        return results

    def export_break(self, br):
        pass

    def export_run(self, run):
        if self.first_pass:
            if self.captured_runs is not None:
                self.captured_runs.append(run)

        # TODO squash multiple sequential text nodes into one?
        results = self.yield_nested(run.children, self.export_node)
        if run.effective_properties:
            results = self.export_run_apply_properties(run, results)
        return results

    def get_run_styles_to_apply(self, run):
        properties = run.effective_properties
        property_rules = [
            (properties.bold, self.export_run_property_bold),
            (properties.italic, self.export_run_property_italic),
            (properties.underline, self.export_run_property_underline),
            (properties.caps, self.export_run_property_caps),
            (properties.small_caps, self.export_run_property_small_caps),
            (properties.strike, self.export_run_property_strike),
            (properties.dstrike, self.export_run_property_dstrike),
            (properties.vanish, self.export_run_property_vanish),
            (properties.hidden, self.export_run_property_hidden),
            (properties.vertical_align, self.export_run_property_vertical_align),
        ]
        for actual_value, handler in property_rules:
            if actual_value:
                yield handler

    def export_run_apply_properties(self, run, results):
        styles_to_apply = self.get_run_styles_to_apply(run)
        # Preserve style application order, but don't duplicate
        applied_styles = set()
        for func in styles_to_apply:
            if func not in applied_styles:
                results = func(run, results)
                applied_styles.add(func)
        return results

    def export_run_property_bold(self, run, results):
        return results

    def export_run_property_italic(self, run, results):
        return results

    def export_run_property_underline(self, run, results):
        return results

    def export_run_property_caps(self, run, results):
        return results

    def export_run_property_small_caps(self, run, results):
        return results

    def export_run_property_strike(self, run, results):
        return results

    def export_run_property_dstrike(self, run, results):
        return results

    def export_run_property_vanish(self, run, results):
        return results

    def export_run_property_hidden(self, run, results):
        return results

    def export_run_property_vertical_align(self, run, results):
        return results

    def export_hyperlink(self, hyperlink):
        return self.yield_nested(hyperlink.children, self.export_node)

    def export_text(self, text):
        if not text.text:
            yield ''
        else:
            # TODO should we do this here, or in the HTML exporter?
            yield self.escape(text.text)

    def export_deleted_text(self, deleted_text):
        pass

    def export_no_break_hyphen(self, hyphen):
        yield '-'

    def export_table(self, table):
        return self.yield_nested(table.rows, self.export_node)

    def export_table_row(self, table_row):
        return self.yield_nested(table_row.cells, self.export_node)

    def export_table_cell(self, table_cell):
        numbering_spans = self.yield_numbering_spans(table_cell.children)
        return self.yield_nested(numbering_spans, self.export_node)

    def escape(self, text):
        #  TODO should we use escape here instead?
        return xml.sax.saxutils.quoteattr(text)[1:-1]

    def export_drawing(self, drawing):
        pass

    def export_smart_tag_run(self, smart_tag):
        return self.yield_nested(smart_tag.children, self.export_node)

    def export_inserted_run(self, inserted_run):
        return self.yield_nested(inserted_run.children, self.export_node)

    def export_picture(self, picture):
        return self.yield_nested(picture.children, self.export_node)

    def export_deleted_run(self, deleted_run):
        return self.yield_nested(deleted_run.children, self.export_node)

    def export_footnote_reference(self, footnote_reference):
        if self.first_pass:
            return

        if footnote_reference.footnote is None:
            return
        self.footnote_tracker.append(footnote_reference)
        footnote_index = len(self.footnote_tracker)
        yield '{0}'.format(footnote_index)

    def export_footnote(self, footnote):
        # TODO a footnote can have paragraph children, should we track
        # numbering?
        return self.yield_nested(footnote.children, self.export_node)

    def export_footnotes(self):
        if not self.footnote_tracker:
            return
        for footnote_reference in self.footnote_tracker:
            footnote = footnote_reference.footnote
            if footnote:
                for result in self.export_node(footnote):
                    yield result

    def export_footnote_reference_mark(self, footnote_reference_mark):
        pass

    def export_vml_shape(self, shape):
        return self.yield_nested(shape.children, self.export_node)

    def export_object(self, _object):
        return self.yield_nested(_object.children, self.export_node)

    def export_vml_image_data(self, image_data):
        pass

    def export_sdt(self, sdt):
        return self.export_node(sdt.content)

    def export_sdt_content(self, sdt_content):
        return self.yield_nested(sdt_content.children, self.export_node)

    def export_sdt_run(self, sdt_run):
        return self.export_sdt(sdt_run)

    def export_sdt_content_run(self, sdt_content_run):
        return self.export_sdt_content(sdt_content_run)

    def export_sdt_block(self, sdt_block):
        return self.export_sdt(sdt_block)

    def export_sdt_content_block(self, sdt_content_block):
        return self.export_sdt_content(sdt_content_block)

    def export_tab_char(self, tab_char):
        pass

    def export_numbering_span(self, numbering_span):
        return self.yield_nested(numbering_span.children, self.export_node)

    def export_numbering_item(self, numbering_item):
        return self.yield_nested(numbering_item.children, self.export_node)

    def export_simple_field(self, simple_field):
        default_results = self.yield_nested(
            simple_field.children, self.export_node)

        parsed_instr = simple_field.parse_instr()
        if not parsed_instr:
            return default_results

        field_type, field_args = parsed_instr
        func = self.field_type_to_export_func_map.get(field_type, None)
        if callable(func):
            return func(simple_field, field_args)
        return default_results

    def export_field_char(self, field_char):
        if self.first_pass:
            if field_char.is_type_begin():
                self.captured_runs = [field_char.parent]
            elif field_char.is_type_end() and self.captured_runs is not None:
                self.complex_field_runs.extend(self.captured_runs)
                self.captured_runs = None
            return

    def export_field_code(self, field_code):
        pass
