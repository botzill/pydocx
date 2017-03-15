# -*- coding: utf-8 -*-
from __future__ import (
    absolute_import,
    print_function,
    unicode_literals,
)

import sys
import time

from pydocx.openxml.packaging import ImagePart

from pydocx.test import TranslationTestCase
from pydocx.test.document_builder import DocxBuilder as DXB


class ImageLocal(TranslationTestCase):
    relationships = [
        dict(
            relationship_id='rId0',
            relationship_type=ImagePart.relationship_type,
            external=False,
            target_path='media/image1.jpeg',
            data=b'content1',
        ),
        dict(
            relationship_id='rId1',
            relationship_type=ImagePart.relationship_type,
            external=False,
            target_path='media/image2.jpeg',
            data=b'content2',
        ),
        dict(
            relationship_id='rId2',
            relationship_type=ImagePart.relationship_type,
            external=False,
            target_path='media/image3.jpeg',
            data=b'content3',
        ),
    ]

    expected_output = '''
    <p><img src="data:image/jpeg;base64,Y29udGVudDE=" /></p>
    <p><img src="data:image/jpeg;base64,Y29udGVudDI=" /></p>
    <p><img src="data:image/jpeg;base64,Y29udGVudDM=" /></p>
    '''

    def get_xml(self):
        drawing = DXB.drawing(height=None, width=None, r_id='rId0')
        pict = DXB.pict(height=None, width=None, r_id='rId1')
        rect = DXB.rect(height=None, width=None, r_id='rId2')
        tags = [
            drawing,
            pict,
            rect,
        ]
        body = b''
        for el in tags:
            body += el

        return DXB.xml(body)


class ImageTestCase(TranslationTestCase):
    relationships = [
        dict(
            relationship_id='rId0',
            relationship_type=ImagePart.relationship_type,
            external=False,
            target_path='media/image1.jpeg',
            data=b'content1',
        ),
        dict(
            relationship_id='rId1',
            relationship_type=ImagePart.relationship_type,
            external=False,
            target_path='media/image2.jpeg',
            data=b'content2',
        ),
        dict(
            relationship_id='rId2',
            relationship_type=ImagePart.relationship_type,
            external=False,
            target_path='media/image3.jpeg',
            data=b'content3',
        ),
    ]

    expected_output = '''
        <p>
            <img
                height="20px"
                src="data:image/jpeg;base64,Y29udGVudDE="
                width="40px"
            />
        </p>
        <p>
            <img
                height="21pt"
                src="data:image/jpeg;base64,Y29udGVudDI="
                width="41pt"
            />
        </p>
        <p>
            <img
                height="22pt"
                src="data:image/jpeg;base64,Y29udGVudDM="
                width="42pt"
            />
        </p>
    '''

    def get_xml(self):
        drawing = DXB.drawing(height=20, width=40, r_id='rId0')
        pict = DXB.pict(height=21, width=41, r_id='rId1')
        rect = DXB.rect(height=22, width=42, r_id='rId2')
        tags = [
            drawing,
            pict,
            rect,
        ]
        body = b''
        for el in tags:
            body += el

        xml = DXB.xml(body)
        return xml


class TableTag(TranslationTestCase):
    expected_output = '''
        <table border="1">
            <tr>
                <td><p>AAA</p></td>
                <td><p>BBB</p></td>
            </tr>
            <tr>
                <td><p>CCC</p></td>
                <td><p>DDD</p></td>
            </tr>
        </table>
    '''

    def get_xml(self):
        cell1 = DXB.table_cell(paragraph=DXB.p_tag('AAA'))
        cell2 = DXB.table_cell(paragraph=DXB.p_tag('CCC'))
        cell3 = DXB.table_cell(paragraph=DXB.p_tag('BBB'))
        cell4 = DXB.table_cell(paragraph=DXB.p_tag('DDD'))
        rows = [DXB.table_row([cell1, cell3]), DXB.table_row([cell2, cell4])]
        table = DXB.table(rows)
        body = table
        xml = DXB.xml(body)
        return xml


class RowSpanTestCase(TranslationTestCase):

    expected_output = '''
           <table border="1">
            <tr>
                <td rowspan="2">
                    <p>AAA</p>
                </td>
                <td>
                    <p>BBB</p>
                </td>
            </tr>
            <tr>
                <td>
                    <p>CCC</p>
                </td>
            </tr>
        </table>
    '''

    def get_xml(self):
        cell1 = DXB.table_cell(
            paragraph=DXB.p_tag('AAA'), merge=True, merge_continue=False)
        cell2 = DXB.table_cell(
            paragraph=DXB.p_tag(None), merge=False, merge_continue=True)
        cell3 = DXB.table_cell(paragraph=DXB.p_tag('BBB'))
        cell4 = DXB.table_cell(paragraph=DXB.p_tag('CCC'))
        rows = [DXB.table_row([cell1, cell3]), DXB.table_row([cell2, cell4])]
        table = DXB.table(rows)
        body = table
        xml = DXB.xml(body)
        return xml


class NestedTableTag(TranslationTestCase):
    expected_output = '''
        <table border="1">
            <tr>
                <td><p>AAA</p></td>
                <td><p>BBB</p></td>
            </tr>
            <tr>
                <td><p>CCC</p></td>
                <td>
                    <table border="1">
                        <tr>
                            <td><p>DDD</p></td>
                            <td><p>EEE</p></td>
                        </tr>
                        <tr>
                            <td><p>FFF</p></td>
                            <td><p>GGG</p></td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    '''

    def get_xml(self):
        cell1 = DXB.table_cell(paragraph=DXB.p_tag('DDD'))
        cell2 = DXB.table_cell(paragraph=DXB.p_tag('FFF'))
        cell3 = DXB.table_cell(paragraph=DXB.p_tag('EEE'))
        cell4 = DXB.table_cell(paragraph=DXB.p_tag('GGG'))
        rows = [DXB.table_row([cell1, cell3]), DXB.table_row([cell2, cell4])]
        nested_table = DXB.table(rows)
        cell1 = DXB.table_cell(paragraph=DXB.p_tag('AAA'))
        cell2 = DXB.table_cell(paragraph=DXB.p_tag('CCC'))
        cell3 = DXB.table_cell(paragraph=DXB.p_tag('BBB'))
        cell4 = DXB.table_cell(nested_table)
        rows = [DXB.table_row([cell1, cell3]), DXB.table_row([cell2, cell4])]
        table = DXB.table(rows)
        body = table
        xml = DXB.xml(body)
        return xml


class TableWithInvalidTag(TranslationTestCase):
    expected_output = '''
        <table border="1">
            <tr>
                <td><p>AAA</p></td>
                <td><p>BBB</p></td>
            </tr>
            <tr>
                <td></td>
                <td><p>DDD</p></td>
            </tr>
        </table>
    '''

    def get_xml(self):
        cell1 = DXB.table_cell(paragraph=DXB.p_tag('AAA'))
        cell2 = DXB.table_cell('<w:invalidTag>CCC</w:invalidTag>')
        cell3 = DXB.table_cell(paragraph=DXB.p_tag('BBB'))
        cell4 = DXB.table_cell(paragraph=DXB.p_tag('DDD'))
        rows = [DXB.table_row([cell1, cell3]), DXB.table_row([cell2, cell4])]
        table = DXB.table(rows)
        body = table
        xml = DXB.xml(body)
        return xml


class TableWithListAndParagraph(TranslationTestCase):
    expected_output = '''
        <table border="1">
            <tr>
                <td>
                    <ol class="pydocx-list-style-type-decimal">
                        <li><p>AAA</p></li>
                        <li><p>BBB</p></li>
                    </ol>
                    <p>CCC</p>
                    <p>DDD</p>
                </td>
            </tr>
        </table>
    '''

    numbering_dict = {
        '1': {
            '0': 'decimal',
        }
    }

    def get_xml(self):
        li_text = [
            ('AAA', 0, 1),
            ('BBB', 0, 1),
        ]
        lis = b''
        for text, ilvl, numId in li_text:
            lis += DXB.li(text=text, ilvl=ilvl, numId=numId)
        els = [
            lis,
            DXB.p_tag('CCC'),
            DXB.p_tag('DDD'),
        ]
        td = b''
        for el in els:
            td += el
        cell1 = DXB.table_cell(td)
        row = DXB.table_row([cell1])
        table = DXB.table([row])
        body = table
        xml = DXB.xml(body)
        return xml


class TableWithCellBackgroundColor(TranslationTestCase):
    expected_output = '''
        <table border="1">
            <tr>
                <td style="background-color: #FF00FF">
                    <p>AAA</p>
                </td>
                <td>
                    <p>BBB</p>
                </td>
            </tr>
            <tr>
                <td>
                    <p>CCC</p>
                </td>
                <td style="background-color: #000000">
                    <p>
                        <span style="color: #FFFFFF">DDD</span>
                    </p>
                </td>
            </tr>
        </table>
    '''

    def get_xml(self):
        cell1 = DXB.table_cell(paragraph=DXB.p_tag('AAA'), fill_color='FF00FF')
        cell2 = DXB.table_cell(paragraph=DXB.p_tag('BBB'), fill_color='FFFFFF')
        cell3 = DXB.table_cell(paragraph=DXB.p_tag('CCC'), fill_color='auto')
        # for dark color we get white text color
        cell4 = DXB.table_cell(paragraph=DXB.p_tag('DDD'), fill_color='000000')
        rows = [DXB.table_row([cell1, cell2]), DXB.table_row([cell3, cell4])]
        table = DXB.table(rows)
        body = table
        xml = DXB.xml(body)
        return xml


class SimpleListTestCase(TranslationTestCase):
    expected_output = '''
        <ol class="pydocx-list-style-type-lowerLetter">
            <li><p>AAA</p></li>
            <li><p>BBB</p></li>
            <li><p>CCC</p></li>
        </ol>
    '''

    # Ensure its not failing somewhere and falling back to decimal
    numbering_dict = {
        '1': {
            '0': 'lowerLetter',
        }
    }

    def get_xml(self):
        li_text = [
            ('AAA', 0, 1),
            ('BBB', 0, 1),
            ('CCC', 0, 1),
        ]
        lis = b''
        for text, ilvl, numId in li_text:
            lis += DXB.li(text=text, ilvl=ilvl, numId=numId)

        xml = DXB.xml(lis)
        return xml


class SingleListItemTestCase(TranslationTestCase):
    expected_output = '''
        <ol class="pydocx-list-style-type-lowerLetter">
            <li><p>AAA</p></li>
        </ol>
    '''

    # Ensure its not failing somewhere and falling back to decimal
    numbering_dict = {
        '1': {
            '0': 'lowerLetter',
        }
    }

    def get_xml(self):
        li_text = [
            ('AAA', 0, 1),
        ]
        lis = b''
        for text, ilvl, numId in li_text:
            lis += DXB.li(text=text, ilvl=ilvl, numId=numId)

        xml = DXB.xml(lis)
        return xml


class ListWithContinuationTestCase(TranslationTestCase):
    expected_output = '''
        <ol class="pydocx-list-style-type-decimal">
            <li>
                <p>AAA</p>
                <p>BBB</p>
            </li>
            <li>
                <p>CCC</p>
                <table border="1">
                    <tr>
                        <td><p>DDD</p></td>
                        <td><p>EEE</p></td>
                    </tr>
                    <tr>
                        <td><p>FFF</p></td>
                        <td><p>GGG</p></td>
                    </tr>
                </table>
            </li>
            <li><p>HHH</p></li>
        </ol>
    '''

    def get_xml(self):
        cell1 = DXB.table_cell(paragraph=DXB.p_tag('DDD'))
        cell2 = DXB.table_cell(paragraph=DXB.p_tag('FFF'))
        cell3 = DXB.table_cell(paragraph=DXB.p_tag('EEE'))
        cell4 = DXB.table_cell(paragraph=DXB.p_tag('GGG'))
        rows = [DXB.table_row([cell1, cell3]), DXB.table_row([cell2, cell4])]
        table = DXB.table(rows)
        tags = [
            DXB.li(text='AAA', ilvl=0, numId=1),
            DXB.p_tag('BBB'),
            DXB.li(text='CCC', ilvl=0, numId=1),
            table,
            DXB.li(text='HHH', ilvl=0, numId=1),
        ]
        body = b''
        for el in tags:
            body += el

        xml = DXB.xml(body)
        return xml


class ListWithMultipleContinuationTestCase(TranslationTestCase):
    expected_output = '''
        <ol class="pydocx-list-style-type-decimal">
            <li>
                <p>AAA</p>
                <table border="1">
                    <tr>
                        <td><p>BBB</p></td>
                    </tr>
                </table>
                <table border="1">
                    <tr>
                        <td><p>CCC</p></td>
                    </tr>
                </table>
            </li>
            <li><p>DDD</p></li>
        </ol>
    '''

    def get_xml(self):
        cell = DXB.table_cell(paragraph=DXB.p_tag('BBB'))
        row = DXB.table_row([cell])
        table1 = DXB.table([row])
        cell = DXB.table_cell(paragraph=DXB.p_tag('CCC'))
        row = DXB.table_row([cell])
        table2 = DXB.table([row])
        tags = [
            DXB.li(text='AAA', ilvl=0, numId=1),
            table1,
            table2,
            DXB.li(text='DDD', ilvl=0, numId=1),
        ]
        body = b''
        for el in tags:
            body += el

        xml = DXB.xml(body)
        return xml


class MangledIlvlTestCase(TranslationTestCase):
    expected_output = '''
        <ol class="pydocx-list-style-type-lowerLetter">
            <li><p>AAA</p></li>
        </ol>
        <ol class="pydocx-list-style-type-decimal">
            <li><p>BBB</p></li>
        </ol>
        <ol class="pydocx-list-style-type-decimal">
            <li><p>CCC</p></li>
        </ol>
    '''

    def get_xml(self):
        tags = [
            DXB.li(text='AAA', ilvl=0, numId=2),
            DXB.li(text='BBB', ilvl=1, numId=1),
            DXB.li(text='CCC', ilvl=0, numId=1),
        ]
        body = b''
        for el in tags:
            body += el

        xml = DXB.xml(body)
        return xml


class SeperateListsIntoParentListTestCase(TranslationTestCase):
    expected_output = '''
        <ol class="pydocx-list-style-type-lowerLetter">
            <li>
                <p>AAA</p>
                <ol class="pydocx-list-style-type-decimal">
                    <li><p>BBB</p></li>
                    <li><p>CCC</p></li>
                </ol>
            </li>
            <li><p>DDD</p></li>
        </ol>
    '''

    def get_xml(self):
        tags = [
            DXB.li(text='AAA', ilvl=0, numId=2),
            # Because AAA and DDD are part of the same list (same list id)
            # and BBB,CCC are different, these need to be properly formatted
            # into a single list where BBB,CCC are added as nested list to AAA item
            DXB.li(text='BBB', ilvl=0, numId=1),
            DXB.li(text='CCC', ilvl=0, numId=1),
            DXB.li(text='DDD', ilvl=0, numId=2),
        ]
        body = b''
        for el in tags:
            body += el

        xml = DXB.xml(body)
        return xml


class InvalidIlvlOrderTestCase(TranslationTestCase):
    expected_output = '''
        <ol class="pydocx-list-style-type-decimal">
            <li>
                <p>AAA</p>
                <ol class="pydocx-list-style-type-decimal">
                    <li><p>BBB</p></li>
                </ol>
            </li>
            <li><p>CCC</p></li>
        </ol>
    '''

    numbering_dict = {
        '1': {
            '0': 'decimal',
            '1': 'decimal',
            '2': 'decimal',
            '3': 'decimal',
        }
    }

    def get_xml(self):
        tags = [
            # purposefully start at 1 instead of 0
            DXB.li(text='AAA', ilvl=1, numId=1),
            DXB.li(text='BBB', ilvl=3, numId=1),
            DXB.li(text='CCC', ilvl=2, numId=1),
        ]
        body = b''
        for el in tags:
            body += el

        xml = DXB.xml(body)
        return xml


class DeeplyNestedTableTestCase(TranslationTestCase):
    expected_output = ''
    run_expected_output = False

    def get_xml(self):
        paragraph = DXB.p_tag('AAA')

        for _ in range(1000):
            cell = DXB.table_cell(paragraph)
            row = DXB.table_cell([cell])
            table = DXB.table([row])
        body = table
        xml = DXB.xml(body)
        return xml

    def test_performance(self):
        with self.toggle_run_expected_output():
            start_time = time.time()
            try:
                self.assert_expected_output()
            except AssertionError:
                pass
            end_time = time.time()
            total_time = end_time - start_time
            # This finishes in under a second on python 2.7
            expected_time = 3
            if sys.version_info[0] == 3:
                expected_time = 5  # Slower on python 3
            error_message = 'Total time: %s; Expected time: %d' % (
                total_time,
                expected_time,
            )
            assert total_time < expected_time, error_message


class LargeCellTestCase(TranslationTestCase):
    expected_output = ''
    run_expected_output = False

    def get_xml(self):
        # Make sure it is over 1000 (which is the recursion limit)
        paragraphs = [DXB.p_tag('%d' % i) for i in range(1000)]
        cell = DXB.table_cell(paragraphs)
        row = DXB.table_cell([cell])
        table = DXB.table([row])
        body = table
        xml = DXB.xml(body)
        return xml

    def test_performance(self):
        with self.toggle_run_expected_output():
            start_time = time.time()
            try:
                self.assert_expected_output()
            except AssertionError:
                pass
            end_time = time.time()
            total_time = end_time - start_time
            # This finishes in under a second on python 2.7
            expected_time = 3
            if sys.version_info[0] == 3:
                expected_time = 7  # Slower on python 3
            error_message = 'Total time: %s; Expected time: %d' % (
                total_time,
                expected_time,
            )
            assert total_time < expected_time, error_message


class NonStandardTextTagsTestCase(TranslationTestCase):
    expected_output = '''
        <p><span class="pydocx-insert">insert </span>
        smarttag</p>
    '''

    def get_xml(self):
        run_tags = [DXB.r_tag([DXB.t_tag(i)]) for i in 'insert ']
        insert_tag = DXB.insert_tag(run_tags)
        run_tags = [DXB.r_tag([DXB.t_tag(i)]) for i in 'smarttag']
        smart_tag = DXB.smart_tag(run_tags)

        run_tags = [insert_tag, smart_tag]
        body = DXB.p_tag(run_tags)
        xml = DXB.xml(body)
        return xml


class RTagWithNoText(TranslationTestCase):
    expected_output = '<p>&#160;</p>'

    def get_xml(self):
        p_tag = DXB.p_tag(None)  # No text
        run_tags = [p_tag]
        # The bug is only present in a hyperlink
        run_tags = [DXB.hyperlink_tag(r_id='rId0', run_tags=run_tags)]
        body = DXB.p_tag(run_tags)

        xml = DXB.xml(body)
        return xml


class DeleteTagInList(TranslationTestCase):
    expected_output = '''
        <ol class="pydocx-list-style-type-decimal">
            <li>
                <p>AAA</p>
                <p>
                    <span class="pydocx-delete">BBB</span>
                </p>
            </li>
            <li><p>CCC</p></li>
        </ol>
    '''

    numbering_dict = {
        '1': {
            '0': 'decimal',
        }
    }

    def get_xml(self):
        delete_tags = DXB.delete_tag(['BBB'])
        p_tag = DXB.p_tag([delete_tags])

        body = DXB.li(text='AAA', ilvl=0, numId=1)
        body += p_tag
        body += DXB.li(text='CCC', ilvl=0, numId=1)

        xml = DXB.xml(body)
        return xml


class InsertTagInList(TranslationTestCase):
    expected_output = '''
        <ol class="pydocx-list-style-type-decimal">
            <li>
                <p>AAA</p>
                <p>
                    <span class="pydocx-insert">BBB</span>
                </p>
            </li>
            <li><p>CCC</p></li>
        </ol>
    '''

    numbering_dict = {
        '1': {
            '0': 'decimal',
        }
    }

    def get_xml(self):
        run_tags = [DXB.r_tag([DXB.t_tag(i)]) for i in 'BBB']
        insert_tags = DXB.insert_tag(run_tags)
        p_tag = DXB.p_tag([insert_tags])

        body = DXB.li(text='AAA', ilvl=0, numId=1)
        body += p_tag
        body += DXB.li(text='CCC', ilvl=0, numId=1)

        xml = DXB.xml(body)
        return xml


class SmartTagInList(TranslationTestCase):
    expected_output = '''
        <ol class="pydocx-list-style-type-decimal">
            <li>
                <p>AAA</p>
                <p>BBB</p>
            </li>
            <li><p>CCC</p></li>
        </ol>
    '''

    numbering_dict = {
        '1': {
            '0': 'decimal',
        }
    }

    def get_xml(self):
        run_tags = [DXB.r_tag([DXB.t_tag(i)]) for i in 'BBB']
        smart_tag = DXB.smart_tag(run_tags)
        p_tag = DXB.p_tag([smart_tag])

        body = DXB.li(text='AAA', ilvl=0, numId=1)
        body += p_tag
        body += DXB.li(text='CCC', ilvl=0, numId=1)

        xml = DXB.xml(body)
        return xml


class SingleListItem(TranslationTestCase):
    expected_output = '''
        <ol class="pydocx-list-style-type-lowerLetter">
            <li><p>AAA</p></li>
        </ol>
        <p>BBB</p>
    '''

    numbering_dict = {
        '1': {
            '0': 'lowerLetter',
        }
    }

    def get_xml(self):
        li = DXB.li(text='AAA', ilvl=0, numId=1)
        p_tags = [
            DXB.p_tag('BBB'),
        ]
        body = li
        for p_tag in p_tags:
            body += p_tag
        xml = DXB.xml(body)
        return xml


class SimpleTableTest(TranslationTestCase):
    expected_output = '''
        <table border="1">
            <tr>
                <td><p>Blank</p></td>
                <td><p>Column 1</p></td>
                <td><p>Column 2</p></td>
            </tr>
            <tr>
                <td><p>Row 1</p></td>
                <td><p>First</p></td>
                <td><p>Second</p></td>
            </tr>
            <tr>
                <td><p>Row 2</p></td>
                <td><p>Third</p></td>
                <td><p>Fourth</p></td>
            </tr>
        </table>'''

    def get_xml(self):
        cell1 = DXB.table_cell(paragraph=DXB.p_tag('Blank'))
        cell2 = DXB.table_cell(paragraph=DXB.p_tag('Row 1'))
        cell3 = DXB.table_cell(paragraph=DXB.p_tag('Row 2'))
        cell4 = DXB.table_cell(paragraph=DXB.p_tag('Column 1'))
        cell5 = DXB.table_cell(paragraph=DXB.p_tag('First'))
        cell6 = DXB.table_cell(paragraph=DXB.p_tag('Third'))
        cell7 = DXB.table_cell(paragraph=DXB.p_tag('Column 2'))
        cell8 = DXB.table_cell(paragraph=DXB.p_tag('Second'))
        cell9 = DXB.table_cell(paragraph=DXB.p_tag('Fourth'))
        rows = [DXB.table_row([cell1, cell4, cell7]),
                DXB.table_row([cell2, cell5, cell8]),
                DXB.table_row([cell3, cell6, cell9])]
        table = DXB.table(rows)
        body = table
        xml = DXB.xml(body)
        return xml


class MissingIlvl(TranslationTestCase):
    expected_output = '''
        <ol class="pydocx-list-style-type-decimal">
            <li>
                <p>AAA</p>
                <p>BBB</p>
            </li>
            <li><p>CCC</p></li>
        </ol>
    '''

    def get_xml(self):
        li_text = [
            ('AAA', 0, 1),
            ('BBB', None, 1),  # Because why not.
            ('CCC', 0, 1),
        ]
        lis = b''
        for text, ilvl, numId in li_text:
            lis += DXB.li(text=text, ilvl=ilvl, numId=numId)
        body = lis
        xml = DXB.xml(body)
        return xml


class SameNumIdInTable(TranslationTestCase):
    expected_output = '''
        <ol class="pydocx-list-style-type-lowerLetter">
            <li>
                <p>AAA</p>
                <table border="1">
                    <tr>
                        <td>
                            <ol class="pydocx-list-style-type-lowerLetter">
                                <li><p>BBB</p></li>
                            </ol>
                        </td>
                    </tr>
                </table>
            </li>
            <li><p>CCC</p></li>
        </ol>
    '''

    # Ensure its not failing somewhere and falling back to decimal
    numbering_dict = {
        '1': {
            '0': 'lowerLetter',
        }
    }

    def get_xml(self):
        li_text = [
            ('BBB', 0, 1),
        ]
        lis = b''
        for text, ilvl, numId in li_text:
            lis += DXB.li(text=text, ilvl=ilvl, numId=numId)
        cell1 = DXB.table_cell(lis)
        rows = DXB.table_row([cell1])
        table = DXB.table([rows])
        lis = b''
        lis += DXB.li(text='AAA', ilvl=0, numId=1)
        lis += table
        lis += DXB.li(text='CCC', ilvl=0, numId=1)
        body = lis
        xml = DXB.xml(body)
        return xml


class SDTTestCase(TranslationTestCase):
    expected_output = '''
        <ol class="pydocx-list-style-type-decimal">
            <li>
                <p>AAA</p>
                <p>BBB</p>
            </li>
            <li><p>CCC</p></li>
        </ol>
    '''

    numbering_dict = {
        '1': {
            '0': 'decimal',
        }
    }

    def get_xml(self):
        body = b''
        body += DXB.li(text='AAA', ilvl=0, numId=1)
        body += DXB.sdt_tag(p_tag=DXB.p_tag(text='BBB'))
        body += DXB.li(text='CCC', ilvl=0, numId=1)

        xml = DXB.xml(body)
        return xml


class SuperAndSubScripts(TranslationTestCase):
    expected_output = '''
        <p>AAA<sup>BBB</sup></p>
        <p><sub>CCC</sub>DDD</p>
    '''

    def get_xml(self):
        p_tags = [
            DXB.p_tag(
                [
                    DXB.r_tag([DXB.t_tag('AAA')]),
                    DXB.r_tag(
                        [DXB.t_tag('BBB')],
                        rpr=DXB.rpr_tag({'vertAlign': 'superscript'}),
                    ),
                ],
            ),
            DXB.p_tag(
                [
                    DXB.r_tag(
                        [DXB.t_tag('CCC')],
                        rpr=DXB.rpr_tag({'vertAlign': 'subscript'}),
                    ),
                    DXB.r_tag([DXB.t_tag('DDD')]),
                ],
            ),
        ]
        body = b''
        for p_tag in p_tags:
            body += p_tag

        xml = DXB.xml(body)
        return xml


class AvailableInlineTags(TranslationTestCase):
    expected_output = '''
        <p><strong>aaa</strong></p>
        <p><span class="pydocx-underline">bbb</span></p>
        <p><em>ccc</em></p>
        <p><span class="pydocx-caps">ddd</span></p>
        <p><span class="pydocx-small-caps">eee</span></p>
        <p><span class="pydocx-strike">fff</span></p>
        <p><span class="pydocx-strike">ggg</span></p>
        <p><span class="pydocx-hidden">hhh</span></p>
        <p><span class="pydocx-hidden">iii</span></p>
        <p><sup>jjj</sup></p>
    '''

    def get_xml(self):
        p_tags = [
            DXB.p_tag(
                [
                    DXB.r_tag(
                        [DXB.t_tag('aaa')],
                        rpr=DXB.rpr_tag({'b': None}),
                    ),
                ],
            ),
            DXB.p_tag(
                [
                    DXB.r_tag(
                        [DXB.t_tag('bbb')],
                        rpr=DXB.rpr_tag({'u': 'single'}),
                    ),
                ],
            ),
            DXB.p_tag(
                [
                    DXB.r_tag(
                        [DXB.t_tag('ccc')],
                        rpr=DXB.rpr_tag({'i': None}),
                    ),
                ],
            ),
            DXB.p_tag(
                [
                    DXB.r_tag(
                        [DXB.t_tag('ddd')],
                        rpr=DXB.rpr_tag({'caps': None}),
                    ),
                ],
            ),
            DXB.p_tag(
                [
                    DXB.r_tag(
                        [DXB.t_tag('eee')],
                        rpr=DXB.rpr_tag({'smallCaps': None}),
                    ),
                ],
            ),
            DXB.p_tag(
                [
                    DXB.r_tag(
                        [DXB.t_tag('fff')],
                        rpr=DXB.rpr_tag({'strike': None})
                    ),
                ],
            ),
            DXB.p_tag(
                [
                    DXB.r_tag(
                        [DXB.t_tag('ggg')],
                        rpr=DXB.rpr_tag({'dstrike': None}),
                    ),
                ],
            ),
            DXB.p_tag(
                [
                    DXB.r_tag(
                        [DXB.t_tag('hhh')],
                        rpr=DXB.rpr_tag({'vanish': None})
                    ),
                ],
            ),
            DXB.p_tag(
                [
                    DXB.r_tag(
                        [DXB.t_tag('iii')],
                        rpr=DXB.rpr_tag({'webHidden': None}),
                    ),
                ],
            ),
            DXB.p_tag(
                [
                    DXB.r_tag(
                        [DXB.t_tag('jjj')],
                        rpr=DXB.rpr_tag({'vertAlign': 'superscript'}),
                    ),
                ],
            ),
        ]
        body = b''
        for p_tag in p_tags:
            body += p_tag

        xml = DXB.xml(body)
        return xml


class NestedListTestCase(TranslationTestCase):
    expected_output = u"""
    <ol class="pydocx-list-style-type-decimal">
        <li>
            <p>AAA</p>
            <ol class="pydocx-list-style-type-decimal">
                <li>
                    <p>BBB</p>
                </li>
            </ol>
        </li>
        <li>
            <p>CCC</p>
            <ol class="pydocx-list-style-type-decimal">
                <li>
                    <p>DDD</p>
                    <ol class="pydocx-list-style-type-decimal">
                        <li>
                            <p>EEE</p>
                        </li>
                    </ol>
                </li>
            </ol>
        </li>
    </ol>
    """

    numbering_dict = {
        '1': {
            '0': 'decimal',
            '1': 'decimal',
            '2': 'decimal',
        },
    }

    def get_xml(self):
        li_text = [
            ('AAA', 0, 1),
            ('BBB', 1, 1),
            ('CCC', 0, 1),
            ('DDD', 1, 1),
            ('EEE', 2, 1),
        ]
        lis = b''
        for text, ilvl, numId in li_text:
            lis += DXB.li(text=text, ilvl=ilvl, numId=numId)

        xml = DXB.xml(lis)
        return xml


class MultipleNestedListTestCase(TranslationTestCase):
    expected_output = u"""
    <ol class="pydocx-list-style-type-decimal">
        <li>
            <p>AAA</p>
            <ol class="pydocx-list-style-type-decimal">
                <li>
                    <p>BBB</p>
                    <ol class="pydocx-list-style-type-decimal">
                        <li><p>CCC</p></li>
                        <li><p>DDD</p></li>
                    </ol>
                </li>
                <li>
                    <p>EEE</p>
                    <ol class="pydocx-list-style-type-decimal">
                        <li><p>FFF</p></li>
                        <li><p>GGG</p></li>
                    </ol>
                </li>
                <li>
                    <p>HHH</p>
                    <ol class="pydocx-list-style-type-decimal">
                        <li><p>III</p></li>
                        <li><p>JJJ</p></li>
                    </ol>
                </li>
            </ol>
        </li>
        <li><p>KKK</p></li>
    </ol>
    <ol class="pydocx-list-style-type-lowerLetter">
        <li>
            <p>LLL</p>
            <ol class="pydocx-list-style-type-lowerLetter">
                <li><p>MMM</p></li>
                <li><p>NNN</p></li>
            </ol>
        </li>
        <li>
            <p>OOO</p>
            <ol class="pydocx-list-style-type-lowerLetter">
                <li><p>PPP</p></li>
                <li>
                    <p>QQQ</p>
                    <ol class="pydocx-list-style-type-decimal">
                        <li><p>RRR</p></li>
                    </ol>
                </li>
            </ol>
        </li>
        <li>
            <p>SSS</p>
            <ol class="pydocx-list-style-type-lowerLetter">
                <li><p>TTT</p></li>
                <li><p>UUU</p></li>
            </ol>
        </li>
    </ol>
    """

    numbering_dict = {
        '1': {
            '0': 'decimal',
            '1': 'decimal',
            '2': 'decimal',
        },
        '2': {
            '0': 'lowerLetter',
            '1': 'lowerLetter',
            '2': 'decimal',
        },
    }

    def get_xml(self):
        li_text = [
            ('AAA', 0, 1),
            ('BBB', 1, 1),
            ('CCC', 2, 1),
            ('DDD', 2, 1),
            ('EEE', 1, 1),
            ('FFF', 2, 1),
            ('GGG', 2, 1),
            ('HHH', 1, 1),
            ('III', 2, 1),
            ('JJJ', 2, 1),
            ('KKK', 0, 1),
            ('LLL', 0, 2),
            ('MMM', 1, 2),
            ('NNN', 1, 2),
            ('OOO', 0, 2),
            ('PPP', 1, 2),
            ('QQQ', 1, 2),
            ('RRR', 2, 2),
            ('SSS', 0, 2),
            ('TTT', 1, 2),
            ('UUU', 1, 2),
        ]
        lis = b''
        for text, ilvl, numId in li_text:
            lis += DXB.li(text=text, ilvl=ilvl, numId=numId)

        xml = DXB.xml(lis)
        return xml


class ParagraphBordersTestCase(TranslationTestCase):
    expected_output = '''
        <div style="border-top:1pt solid #000000;padding:0pt">
            <p>AAA</p>
        </div>
        <div style="border-bottom:0.6pt solid #0000FF;padding:0pt 0pt 2pt 0pt">
            <p>BBB</p>
        </div>
    '''

    def get_xml(self):
        p1_border = {
            'top': {
                'color': 'auto'
            }
        }
        p2_border = {
            'bottom': {
                'val': 'single',
                'sz': '5',
                'space': '2',
                'color': '0000FF'
            }
        }
        p_tags = [
            DXB.p_tag('AAA', borders=p1_border),
            DXB.p_tag('BBB', borders=p2_border)
        ]
        body = b''
        for p_tag in p_tags:
            body += p_tag
        xml = DXB.xml(body)
        return xml


class ParagraphBordersSameTopBottomTestCase(TranslationTestCase):
    expected_output = '''
        <div style="border:0.6pt solid #0000FF;padding:2pt">
            <p>AAA</p>
        </div>
        <div style="border-top:0.6pt solid #0000FF;padding:2pt 0pt 0pt 0pt">
            <p>BBB</p>
        </div>
    '''

    def get_xml(self):
        p1_border = {
            'top': {
                'val': 'single',
                'sz': '5',
                'space': '2',
                'color': '0000FF'
            },
            'bottom': {
                'val': 'single',
                'sz': '5',
                'space': '2',
                'color': '0000FF'
            },
            'left': {
                'val': 'single',
                'sz': '5',
                'space': '2',
                'color': '0000FF'
            },
            'right': {
                'val': 'single',
                'sz': '5',
                'space': '2',
                'color': '0000FF'
            },
        }
        p2_border = {
            'top': {
                'val': 'single',
                'sz': '5',
                'space': '2',
                'color': '0000FF'
            }
        }

        p_tags = [
            DXB.p_tag('AAA', borders=p1_border),
            DXB.p_tag('BBB', borders=p2_border),
        ]
        body = b''
        for p_tag in p_tags:
            body += p_tag
        xml = DXB.xml(body)
        return xml


class ParagraphBordersSameTopBottomWithShadingTestCase(TranslationTestCase):
    expected_output = '''
    <div style="background-color:#0000FF;border:0.6pt solid #0000FF;padding:2pt">
       <p>AAA</p>
   </div>
   <div style="background-color:#FF00FF;border:0.6pt solid #0000FF;border-top:0;padding:2pt">
      <p>BBB</p>
    </div>
'''

    def get_xml(self):
        p1_border = {
            'top': {
                'val': 'single',
                'sz': '5',
                'space': '2',
                'color': '0000FF'
            },
            'bottom': {
                'val': 'single',
                'sz': '5',
                'space': '2',
                'color': '0000FF'
            },
            'left': {
                'val': 'single',
                'sz': '5',
                'space': '2',
                'color': '0000FF'
            },
            'right': {
                'val': 'single',
                'sz': '5',
                'space': '2',
                'color': '0000FF'
            },
        }
        p1_shading = {
            'fill': '0000FF',
        }
        p2_border = p1_border.copy()
        p2_shading = {
            'fill': 'FF00FF',
        }

        p_tags = [
            DXB.p_tag('AAA', borders=p1_border, shading=p1_shading),
            DXB.p_tag('BBB', borders=p2_border, shading=p2_shading),
        ]
        body = b''
        for p_tag in p_tags:
            body += p_tag
        xml = DXB.xml(body)
        return xml


class ParagraphShadingTestCase(TranslationTestCase):
    expected_output = '''
        <div style="background-color:#0000FF">
            <p>AAA</p>
            <p>BBB</p>
        </div>
        <div style="background-color:#000000">
            <p>CCC</p>
        </div>
        <div style="background-color:#FF00FF">
            <p>DDD</p>
        </div>
    '''

    def get_xml(self):
        # <w:shd w:val="clear" w:color="auto" w:fill="92D050"/>
        p1_shading = {
            'fill': '0000FF',
        }
        p2_shading = {
            'fill': '0000FF',
        }
        p3_shading = {
            'val': 'solid',
            'color': 'auto'
        }
        p4_shading = {
            'val': 'solid',
            'color': 'FF00FF'
        }

        p_tags = [
            DXB.p_tag('AAA', shading=p1_shading),
            DXB.p_tag('BBB', shading=p2_shading),
            DXB.p_tag('CCC', shading=p3_shading),
            DXB.p_tag('DDD', shading=p4_shading),
        ]
        body = b''
        for p_tag in p_tags:
            body += p_tag
        xml = DXB.xml(body)
        return xml


class ParagraphBorderAndShadingTestCase(TranslationTestCase):
    expected_output = '''
        <div style="background-color:#0000FF;border-top:0.6pt solid #0000FF;padding:0pt">
            <p>AAA</p>
            <p>BBB</p>
        </div>
        <div style="background-color:#FF00FF;border-bottom:0.6pt solid #0000FF;padding:0pt">
            <p>CCC</p>
        </div>
        <div style="background-color:#FFFFFF;border-bottom:0.6pt solid #0000FF;padding:0pt">
            <p>DDD</p>
        </div>
    '''

    def get_xml(self):
        p1_border = {
            'top': {
                'val': 'single',
                'sz': '5',
                'space': '0',
                'color': '0000FF'
            }
        }
        p1_shading = {
            'fill': '0000FF',
        }
        p2_border = {
            'top': {
                'val': 'single',
                'sz': '5',
                'space': '0',
                'color': '0000FF'
            }
        }
        p2_shading = {
            'fill': '0000FF',
        }
        p3_border = {
            'bottom': {
                'val': 'single',
                'sz': '5',
                'space': '0',
                'color': '0000FF'
            }
        }
        p3_shading = {
            'val': 'solid',
            'color': 'FF00FF',
            'fill': '000000',
        }

        p4_border = {
            'bottom': {
                'val': 'single',
                'sz': '5',
                'space': '0',
                'color': '0000FF'
            }
        }
        p4_shading = {
            'val': 'solid',
            'color': 'FFFFFF',
            'fill': '000000',
        }

        p_tags = [
            DXB.p_tag('AAA', borders=p1_border, shading=p1_shading),
            DXB.p_tag('BBB', borders=p2_border, shading=p2_shading),
            DXB.p_tag('CCC', borders=p3_border, shading=p3_shading),
            DXB.p_tag('DDD', borders=p4_border, shading=p4_shading),
        ]
        body = b''
        for p_tag in p_tags:
            body += p_tag
        xml = DXB.xml(body)
        return xml


class RunBordersTestCase(TranslationTestCase):
    expected_output = '''
        <p>
            <span style="border:1pt solid #00FF00">AAA</span>
            <span style="border:1pt solid #00FF00;padding:1pt">BBB</span>
        </p>
    '''

    def get_xml(self):
        r1_border = {
            'bdr': {
                'color': '00FF00',
                'space': '0'
            }
        }
        r2_border = {
            'bdr': {
                'color': '00FF00',
                'space': '1'
            }
        }

        p_tags = [
            DXB.p_tag(
                [
                    DXB.r_tag(
                        [DXB.t_tag('AAA')],
                        borders=r1_border
                    ),
                    DXB.r_tag(
                        [DXB.t_tag('BBB')],
                        borders=r2_border
                    ),
                ],
            )
        ]
        body = b''
        for p_tag in p_tags:
            body += p_tag

        xml = DXB.xml(body)
        return xml


class RunShadingTestCase(TranslationTestCase):
    expected_output = '''
        <p>
            <span style="background-color:#0000FF">AAABBB</span>
            <span style="background-color:#000000">CCC</span>
            <span style="background-color:#FF00FF">DDD</span>
        </p>
    '''

    def get_xml(self):
        r1_shading = {
            'fill': '0000FF',
        }
        r2_shading = {
            'fill': '0000FF',
        }
        r3_shading = {
            'val': 'solid',
            'color': 'auto'
        }
        r4_shading = {
            'val': 'solid',
            'color': 'FF00FF'
        }

        p_tags = [
            DXB.p_tag(
                [
                    DXB.r_tag(
                        [DXB.t_tag('AAA')],
                        shading=r1_shading
                    ),
                    DXB.r_tag(
                        [DXB.t_tag('BBB')],
                        shading=r2_shading
                    ),
                    DXB.r_tag(
                        [DXB.t_tag('CCC')],
                        shading=r3_shading
                    ),
                    DXB.r_tag(
                        [DXB.t_tag('DDD')],
                        shading=r4_shading
                    ),
                ],
            )
        ]
        body = b''
        for p_tag in p_tags:
            body += p_tag
        xml = DXB.xml(body)
        return xml


class RunBordersAndShadingTestCase(TranslationTestCase):
    expected_output = '''
<p>
    <span style="background-color:#0000FF;border:1pt solid #00FF00">AAABBB</span>
    <span style="background-color:#000000;border:1pt solid #FFFF00;padding:1pt">CCC</span>
    <span style="background-color:#FF00FF;border:1pt solid #FFFF00;padding:1pt">DDD</span>
</p>
'''

    def get_xml(self):
        r1_border = {
            'bdr': {
                'color': '00FF00',
                'space': '0'
            }
        }
        r1_shading = {
            'fill': '0000FF',
        }
        r2_border = {
            'bdr': {
                'color': '00FF00',
                'space': '0'
            }
        }
        r2_shading = {
            'fill': '0000FF',
        }
        r3_border = {
            'bdr': {
                'color': 'FFFF00',
                'space': '1'
            }
        }
        r3_shading = {
            'val': 'solid',
            'color': 'auto'
        }
        r4_border = {
            'bdr': {
                'color': 'FFFF00',
                'space': '1'
            }
        }
        r4_shading = {
            'val': 'solid',
            'color': 'FF00FF'
        }

        p_tags = [
            DXB.p_tag(
                [
                    DXB.r_tag(
                        [DXB.t_tag('AAA')],
                        borders=r1_border,
                        shading=r1_shading
                    ),
                    DXB.r_tag(
                        [DXB.t_tag('BBB')],
                        borders=r2_border,
                        shading=r2_shading
                    ),
                    DXB.r_tag(
                        [DXB.t_tag('CCC')],
                        borders=r3_border,
                        shading=r3_shading
                    ),
                    DXB.r_tag(
                        [DXB.t_tag('DDD')],
                        borders=r4_border,
                        shading=r4_shading
                    ),
                ],
            )
        ]
        body = b''
        for p_tag in p_tags:
            body += p_tag
        xml = DXB.xml(body)
        return xml
