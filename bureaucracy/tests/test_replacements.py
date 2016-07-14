import os
import re
from copy import copy
from zipfile import ZipFile

from bureaucracy import HTML, Image, Table
from bureaucracy.tests.test_fields import DocxTestsBase, resources_dir
from docx.enum.style import WD_STYLE_TYPE


class ImageReplacementTests(DocxTestsBase):
    def setUp(self):
        self.generated_path = os.path.join(resources_dir, 'generated_image.docx')

    def tearDown(self):
        if os.path.exists(self.generated_path):
            os.remove(self.generated_path)

    def test_replace_image(self):
        doc = copy(self._get_docx('image'))

        doc.replace_fields({
            'image': Image(os.path.join(resources_dir, 'pigeon.jpg')),
        })

        doc.save(self.generated_path)

        self.assertEqual(len(doc.get_field_names()), 0)
        self.assertEqual(len(doc._element.xpath('w:body/w:p/w:r/w:drawing')), 1)

        self.assertTrue(any(re.match('word/media/.*\.jpg', f.filename)
                            for f in ZipFile(self.generated_path).infolist()))


class TableReplacement(DocxTestsBase):
    def test_replace_table(self):
        doc = self._get_docx('table')

        doc.replace_fields({
            'table': Table([[str(i * j) for j in range(1, 5)] for i in range(1, 3)],
                           headers=['one', 'two', 'three', 'four']),
        })

        self.assertEqual(len(doc.get_field_names()), 0)

        self.assertTrue(doc._element.xpath('.//w:tbl'))
        self.assertEqual(len(doc._element.xpath('.//w:tbl/w:tr')), 3)
        self.assertEqual(len(doc._element.xpath('(.//w:tbl//w:tr)[1]/w:tc')), 4)
        self.assertEqual(doc._element.xpath('((.//w:tbl//w:tr)[1]/w:tc//w:t//text())'), ['one', 'two', 'three', 'four'])
        self.assertEqual(doc._element.xpath('((.//w:tbl//w:tr)[3]/w:tc//w:t//text())'), ['2', '4', '6', '8'])

    def test_replace_table_no_headers(self):
        doc = self._get_docx('table')

        doc.replace_fields({
            'table': Table([[str(i * j) for j in range(1, 5)] for i in range(1, 3)], headers=None)
        })

        self.assertEqual(len(doc.get_field_names()), 0)

        self.assertTrue(doc._element.xpath('.//w:tbl'))
        self.assertEqual(len(doc._element.xpath('.//w:tbl/w:tr')), 2)
        self.assertEqual(len(doc._element.xpath('(.//w:tbl//w:tr)[1]/w:tc')), 4)
        self.assertEqual(doc._element.xpath('((.//w:tbl//w:tr)[2]/w:tc//w:t//text())'), ['2', '4', '6', '8'])

    def test_malformed_table(self):
        doc = self._get_docx('table')

        rows = [[str(i * j) for j in range(1, 5)] for i in range(1, 3)]
        rows.append(['this row is not the same length'])

        with self.assertRaises(ValueError) as cm:
            doc.replace_fields({
                'table': Table(rows,
                               headers=['one', 'two', 'three', 'four']),
            })

            self.assertEqual(cm.exception.message, 'Data contains rows of varying sizes')

    def test_malformed_headers(self):
        doc = self._get_docx('table')

        rows = [[str(i * j) for j in range(1, 5)] for i in range(1, 3)]
        rows.append(['this row is not the same length'])

        with self.assertRaises(ValueError) as cm:
            doc.replace_fields({
                'table': Table(rows,
                               headers=['one', 'two', 'three']),
            })

            self.assertEqual(cm.exception.message, 'Data and header lengths do not match')


class HTMLReplacementTests(DocxTestsBase):
    def test_html_replacement(self):
        doc = self._get_docx('html')

        html = """
        <h1>Header</h1>

        <h2>Subheader</h2>

        <p><strong>bold</strong> - not bold</p>

        <ul>
            <li>hop</li>
            <li>la</li>
            <li>kee</li>
        </ul>"""

        doc.replace_fields({
            'html': HTML(html)
        })

        self.assertEqual(len(doc.get_field_names()), 0)

        self.assertTrue(doc._element.xpath('.//text()="Subheader"'))
        self.assertFalse(doc._element.xpath('.//text()="h1"'))

        # the run with the "bold" text in it has an rPr tag with a b tag in it
        self.assertTrue(doc._element.xpath('.//*[text()="bold"]/parent::w:r/w:rPr/w:b'))

        # the run with the "not bold" text in it exists but does not have a rPr with a b tag in it
        self.assertTrue(doc._element.xpath('.//*[contains(text(), "not bold")]/parent::*'))
        self.assertFalse(doc._element.xpath('.//*[contains(text(), "not bold")]/parent::*/w:b'))

        # assert that all styles mentioned in the document are also in the document's style section
        style_ids = doc.element.xpath('.//w:pStyle/@w:val')
        self.assertTrue(all(doc.styles.get_style_id(sid, WD_STYLE_TYPE.PARAGRAPH) for sid in style_ids))
