import itertools
import logging
import os
import re
import shutil
import subprocess
import tempfile
from copy import deepcopy
from io import BytesIO

from docx.document import Document
from docx.opc.constants import CONTENT_TYPE
from docx.package import Package
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from lxml.etree import tostring

from bureaucracy.replacements import (HTMLReplacement, ImageReplacement,
                                      Replacement, TableReplacement,
                                      TextReplacement)
from bureaucracy.utils import namespaced

r = re.compile(r' MERGEFIELD +"?([^ ]+?)"? +(|\\\* MERGEFORMAT )', re.I)  # fixme. it might be not a simple as that
# the ooxml standard says that the instrText can be divided over several runs. i have not been able to make microsoft
# word generate something like this though...

logger = logging.getLogger('bureaucracy')


class DocxTemplate(Document):
    def __init__(self, docx, strict=False):
        """
        Initialize the DocxTemplate with the given docx file or stream.
        :param docx: A string or file like object (django.core.File objects work) representing a docx file.
        :param strict: will make the template render throw an error when fields and context do not match if True
        """
        self.strict = strict
        document_part = Package.open(docx).main_document_part
        if document_part.content_type != CONTENT_TYPE.WML_DOCUMENT_MAIN:
            tmpl = "file '%s' is not a Word file, content type is '%s'"
            raise ValueError(tmpl % (docx, document_part.content_type))
        super().__init__(document_part._element, document_part)

    def get_field_names(self):
        """
        Get the name of the mailmerge fields included in the document
        :return: a set of the fields' names
        """
        return {rv[0] for rv in self.iter_fields()}

    def iter_fields(self):
        """
        Gernerator for fldSimple and instrText fields and their fieldnames
        :return: a generator yielding tuples (field name, field)-tuples.
        """
        # there are two ways a mergefield can be represented: the simple way with fldSimple and the more
        # complex way with instrText and fldChar.

        simple_fields = ((field.attrib[namespaced('instr')], field) for field in self._element.xpath('.//w:fldSimple'))
        complex_fields = ((field.text, field) for field in self._element.xpath('.//w:instrText'))

        for instr, field in itertools.chain(simple_fields, complex_fields):

            m = r.match(instr)
            if not m and self.strict:
                raise ValueError("Could not determine name of merge field with instr text '{}'".format(instr))
            elif not m:
                logger.warning("Could not determine name of merge field with instr text '{}'. Skipping".format(instr))
                continue

            yield m.group(1), field

    def replace_simple_field(self, field, replacement):

        # a fldSimple tag is easily replaced, we just create a new run in the same paragraph and replace that one
        # with the fldSimple node.

        parent_node = field.getparent()

        # the standard says that this is the case most of the time so we only deal with this case for now:
        assert parent_node.tag == namespaced('p')

        current_paragraph = Paragraph(parent_node, self._body)
        replacement_run = Run(current_paragraph._p._add_r(), current_paragraph)
        parent_node.replace(field, replacement_run._element)
        replacement.fill(replacement_run)

    def replace_complex_field(self, field, replacement):
        # fldChar is more complex. it's not a tag, but rather a series of fldChar and instrText tags inside separate
        # runs. The tags that concern us are these:
        #
        #  1. <w:fldChar w:fldCharType="begin"/>: The beginning of the field
        #  2. <w:instrText xml:space="preserve"> MERGEFIELD test \* MERGEFORMAT </w:instrText> contains the field's name
        #  3. <w:fldChar w:fldCharType="end"/> Marks the end of the field
        #

        # get the run this field is in
        instr_run_node = field.getparent()
        assert instr_run_node.tag == namespaced('r')

        # we now look for the run containing of the opening fldChar for this instrText, which is the first one
        # with an opening fldChar we encounter before the run with instrText
        opening_run_node = instr_run_node
        while not opening_run_node.xpath('w:fldChar[@w:fldCharType="begin"]'):
            opening_run_node = opening_run_node.getprevious()
            if opening_run_node is None:
                raise ValueError(
                    "Could not find beginning of field with instr node '{}'?! Is the document malformed?".format(field))

        # idem for the run containing the closing fldChar, but of course now looking ahead
        closing_run_node = instr_run_node
        while not closing_run_node.xpath('w:fldChar[@w:fldCharType="end"]'):
            closing_run_node = closing_run_node.getnext()
            if closing_run_node is None:
                raise ValueError(
                    "Could not find end of field with instr node '{}'?! Is the document malformed?".format(field))

        # now replace all runs between the opening and closing runs
        current_paragraph = Paragraph(instr_run_node.getparent(), self._body)

        begin_index = current_paragraph._element.index(opening_run_node)
        end_index = current_paragraph._element.index(closing_run_node)

        run = Run(current_paragraph._p._add_r(), current_paragraph)
        current_paragraph._element[begin_index:end_index] = [run._element]

        replacement.fill(run)

    def replace_fields(self, context):
        unused_fields = set()
        unused_values = set(context.keys())

        for field_name, field in self.iter_fields():

            if field_name in context:
                unused_values.discard(field_name)
                replacement = context[field_name]
                if not isinstance(replacement, Replacement):
                    replacement = TextReplacement(replacement)
            else:
                if self.strict:
                    raise ValueError('Could not find field name {} in context'.format(field_name))
                else:
                    replacement = TextReplacement('')
                    unused_fields.add(field_name)

            if field.tag == namespaced('fldSimple'):
                self.replace_simple_field(field, replacement)
            elif field.tag == namespaced('instrText'):
                self.replace_complex_field(field, replacement)

        if unused_fields:
            logger.warn("Fields %s were present in the document, but not in the context. They were removed",
                        unused_fields)
        if unused_values:
            logger.warn(
                "Values %s were present in the context, but no corresponding fields were found in the document.",
                unused_values)

    def render(self, context, format='docx'):
        doc = deepcopy(self)  # take a copy so we can keep using this instance to generate from other contexts
        doc.replace_fields(context)

        if format == 'docx':
            handle = BytesIO()
            doc.save(handle)
            handle.seek(0)
            return handle.read()
        elif format == 'pdf':
            return doc.to_pdf_bytes()
        else:
            raise Exception('Unsupported format.')

    def render_and_save(self, path, context, format='docx'):
        doc = deepcopy(self)  # take a copy so we can keep using this instance to generate from other contexts
        doc.replace_fields(context)

        if format == 'docx':
            doc.save(path)
        elif format == 'pdf':
            doc.to_pdf(path)

    # what follows is a hack.
    # To be able to generate a pdf document/bytes from a merged document, we call libreoffice's headless
    # command line utility to to the generation. The problem is is that that tool only works with files,
    # not with bytestreams or strings, so we save this docx in a tmp dir, throw that into soffice, which
    # generates a pdf file in the same dir with the same name. Then we return bytes dumped into that file or move
    # the file to the desired path, depending on whether we're saving or rendering directly to a stream.
    # All this also makes rendering to pdf very slow. when doing a lot of renders, it might be a good idea to use
    # the batch functionality in the soffice convert tool.

    def _to_pdf(self, path=None):
        _, tmp_doc_path = tempfile.mkstemp(suffix='.docx')
        soffice_outdir = tempfile.gettempdir()
        tmp_pdf_outfilename = '{}.pdf'.format(tmp_doc_path.split(os.sep)[-1].split('.')[0])
        tmp_pdf_path = os.path.join(soffice_outdir, tmp_pdf_outfilename)

        self.save(tmp_doc_path)
        try:
            subprocess.call(['soffice', '--headless',
                             '--convert-to', 'pdf',
                             tmp_doc_path, '--outdir', soffice_outdir],
                            stdout=subprocess.DEVNULL)

            if path:
                shutil.move(tmp_pdf_path, path)
            else:
                with open(tmp_pdf_path, 'rb') as f:
                    return f.read()

        finally:
            if os.path.exists(tmp_doc_path):
                os.remove(tmp_doc_path)

            if os.path.exists(tmp_pdf_path):
                os.remove(tmp_pdf_path)

    def to_pdf(self, path):
        self._to_pdf(path)

    def to_pdf_bytes(self):
        return self._to_pdf()

    def __str__(self):
        return tostring(self._element, pretty_print=True).decode('utf-8')


if __name__ == '__main__':
    examples_dir = os.path.join(os.path.realpath(os.path.dirname(__file__)), '..', 'examples')
    doc = DocxTemplate(os.path.join(examples_dir, 'sample.docx'))
    context = {
        'table': TableReplacement([[str(i * j) for j in range(1, 8)] for i in range(1, 6)],
                                  headers=['one', 'two', 'three', 'four', 'five', 'six', 'seven']),
        'image': ImageReplacement(os.path.join(examples_dir, 'pigeony.jpg')),
        'html': HTMLReplacement(
            open(os.path.join(examples_dir, '..', 'examples', 'test.html')).read()),
        'text': 'some text',
        'hop': 546
    }

    doc.render_and_save(
        os.path.join(examples_dir, 'generated.docx'),
        context)
