import os
import unittest

from bureaucracy import DocxTemplate

resources_dir = os.path.join(os.path.realpath(os.path.dirname(__file__)), 'resources')


class DocxTestsBase(unittest.TestCase):
    def _get_docx(self, name):
        return DocxTemplate('{}/{}.docx'.format(resources_dir, name))


class GetFieldNamesTests(DocxTestsBase):
    def test_get_only_fldChar(self):
        doc = self._get_docx('complex_fields')
        self.assertEqual({'complex', 'complex2'}, doc.get_field_names())

    def test_get_only_fldSimple(self):
        doc = self._get_docx('simple_fields')
        self.assertEqual({'foo', 'bar', 'baz'}, doc.get_field_names())

    def test_get_all_fields(self):
        doc = self._get_docx('simple_and_complex_fields')
        self.assertEqual({'foo', 'bar', 'baz', 'complex', 'complex2', 'simple'}, doc.get_field_names())


class FieldTests(DocxTestsBase):
    def test_replace_only_fldSimple(self):
        doc = self._get_docx('simple_fields')
        doc.replace_fields({'foo': 'frobnicate',
                            'bar': 'determiorate',
                            'baz': 'some kind of string'})

        self.assertEqual(len(doc.get_field_names()), 0)

        self.assertTrue(doc._element.xpath(".//text()='frobnicate'"))
        self.assertTrue(doc._element.xpath(".//text()='determiorate'"))
        self.assertTrue(doc._element.xpath(".//text()='some kind of string'"))

    def test_replace_only_fldChar(self):
        doc = self._get_docx('complex_fields')
        doc.replace_fields({'complex': 'BEEES. AAAAH. BEEEEES',
                            'complex2': 'ಠ_ಠ unifying matrix conventions is the way of the future, Fred.'})

        self.assertEqual(len(doc.get_field_names()), 0)
        self.assertTrue(doc._element.xpath(".//text()='BEEES. AAAAH. BEEEEES'"))
        self.assertTrue(doc._element.xpath(".//text()='ಠ_ಠ unifying matrix conventions is the way of the future, Fred.'"))

    def test_replace_all_fields(self):
        doc = self._get_docx('simple_and_complex_fields')
        doc.replace_fields({'foo': 'frobnicate',
                            'bar': 'determiorate',
                            'baz': 'some kind of string',
                            'complex': 'BEEES. AAAAH. BEEEEES',
                            'complex2': 'ಠ_ಠ unifying matrix conventions is the way of the future, Fred.',
                            'simple': 5.2})

        self.assertEqual(len(doc.get_field_names()), 0)

        self.assertTrue(doc._element.xpath(".//text()='frobnicate'"))
        self.assertTrue(doc._element.xpath(".//text()='determiorate'"))
        self.assertTrue(doc._element.xpath(".//text()='some kind of string'"))
        self.assertTrue(doc._element.xpath(".//text()='BEEES. AAAAH. BEEEEES'"))
        self.assertTrue(doc._element.xpath(".//text()='ಠ_ಠ unifying matrix conventions is the way of the future, Fred.'"))
        self.assertTrue(doc._element.xpath(".//text()='5.2'"))
