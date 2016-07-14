import os
from io import BytesIO

import docx
from bureaucracy.tests.test_fields import DocxTestsBase, resources_dir
from PyPDF2.pdf import PdfFileReader


class RenderTests(DocxTestsBase):
    def setUp(self):
        self.out_path = os.path.join(resources_dir, 'tmp.docx')
        self.out_path_pdf = os.path.join(resources_dir, 'tmp.pdf')


    def tearDown(self):
        if os.path.exists(self.out_path):
            os.remove(self.out_path)

        if os.path.exists(self.out_path_pdf):
            os.remove(self.out_path_pdf)

    def test_save_as_docx(self):
        doc = self._get_docx('complex_fields')
        doc.render_and_save(self.out_path,
                            context={'complex': 'BEEES. AAAAH. BEEEEES',
                                     'complex2': 'ಠ_ಠ unifying matrix conventions is the way of the future, Max.'})

        self.assertTrue(os.path.exists(self.out_path))

        # can the pytohn-docx library parse our result without throwing a hissy fit?
        docx.Document(self.out_path)

    def test_save_as_pdf(self):
        doc = self._get_docx('complex_fields')
        doc.render_and_save(self.out_path_pdf,
                            context={'complex': 'BEEES. AAAAH. BEEEEES',
                                     'complex2': 'ಠ_ಠ unifying matrix conventions is the way of the future, Max.'},
                            format='pdf')

        self.assertTrue(os.path.exists(self.out_path_pdf))

        # does pypdf like our output?
        pdf = open(self.out_path_pdf, "rb")
        PdfFileReader(pdf)
        pdf.close()

    def test_render_to_pdf_bytes(self):
        doc = self._get_docx('complex_fields')
        data = doc.render(context={'complex': 'BEEES. AAAAH. BEEEEES',
                                   'complex2': 'ಠ_ಠ unifying matrix conventions is the way of the future, Max.'},
                          format='pdf')

        # does pypdf like our output?
        PdfFileReader(BytesIO(data))

    def test_render_to_docx_bytes(self):
        doc = self._get_docx('complex_fields')
        data = doc.render(context={'complex': 'BEEES. AAAAH. BEEEEES',
                                            'complex2': 'ಠ_ಠ unifying matrix conventions is the way of the future, Max.'})

        # can the python-docx library parse our result without throwing a hissy fit?
        docx.Document(BytesIO(data))
