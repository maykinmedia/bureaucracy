import os
import tempfile
from copy import copy

import pypandoc
from docx import Document
from docx.oxml import CT_Tbl
from docx.table import Table
from docx.text.paragraph import Paragraph


class Replacement(object):
    def fill(self, el):
        raise NotImplementedError


class RunReplacement(Replacement):
    pass


class ParagraphReplacement(Replacement):
    def fill_paragraph(self, par):
        raise NotImplementedError

    def fill(self, run):
        paragraph = Paragraph(run._element.getparent(), run._parent._parent)
        self.fill_paragraph(paragraph)


class ImageReplacement(RunReplacement):
    def __init__(self, filename, width=None, height=None):
        self.filename = filename
        self.width = width
        self.height = height

    def fill(self, run):
        run.add_picture(self.filename, self.width, self.height)


class HTMLReplacement(ParagraphReplacement):
    def __init__(self, html):
        # fixme: pandoc doesn't allow us to generate a docx stream in memory, so we have to use a temporary file,
        # which makes this part slow in comparison to other replacements. if anyone knows about a good html to
        # docx converter; help yourself.

        _, tmp_file_name = tempfile.mkstemp(suffix='.docx')
        pypandoc.convert(html, format='html', to='docx', outputfile=tmp_file_name)
        doc = Document(tmp_file_name)

        # find all paragraphs in the output
        self.par_nodes = doc._element.xpath('./w:body/w:p')

        # we just add all styles in the output document. This opens us up to the possibility of clashing styles,
        # but for now that doesn't seem to be an issue. only adding the styles used in the output of pandoc and their
        # ancestors in the style hierarchy would be the ideal solutionm, but this works for now.
        self.styles = doc.styles

        os.remove(tmp_file_name)

    def fill_paragraph(self, par):
        body_el = par._element.getparent()
        par_idx = body_el.index(par._element)
        body_el[par_idx:par_idx + 1] = self.par_nodes

        for style in self.styles:
            par.part.document.styles._element.append(copy(style._element))


class TableReplacement(ParagraphReplacement):
    def __init__(self, data, headers):
        if not isinstance(data, (list, tuple)) or any(not isinstance(d, (list, tuple)) for d in data):
            raise Exception("data should be a list or tuple of lists or tuples")
        if headers and not isinstance(headers, (list, tuple)):
            raise Exception("headers should be a list or tuple")

        if isinstance(data, (list, tuple)) and data:
            if any(len(d) != len(data[0]) for d in data):
                raise ValueError("Data contains rows of varying sizes")
            if headers and len(data[0]) != len(headers):
                raise ValueError("Data and header lengths do not match")

        self.headers = headers
        self.data = data

    def fill_paragraph(self, par):

        nr_rows = len(self.data) if self.data else 0
        if self.headers:
            nr_rows += 1

        nr_cols = len(self.data[0]) if self.data else len(self.headers) if self.headers else 0

        # replace the par element with a table xml element and build a docx.Table from it
        # so we can use that to fill it up
        table_el = CT_Tbl.new_tbl(nr_rows, nr_cols, par.part.document._block_width)
        par._element.getparent().replace(par._element, table_el)
        table = Table(table_el, self)

        if self.headers:
            for i, header in enumerate(self.headers):
                table.rows[0].cells[i].text = str(header)

        for row_idx, row_values in enumerate(self.data, start=1 if self.headers else 0):
            for col_idx, cell_value in enumerate(row_values):
                table.rows[row_idx].cells[col_idx].text = str(cell_value)


class TextReplacement(RunReplacement):
    def __init__(self, text):
        self.text = text

    def fill(self, run):
        run.text = str(self.text)
