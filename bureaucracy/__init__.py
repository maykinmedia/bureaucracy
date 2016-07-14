from bureaucracy.replacements import HTMLReplacement, TextReplacement, ImageReplacement, TableReplacement
from bureaucracy.template import DocxTemplate

HTML = HTMLReplacement
Text = TextReplacement
Image = ImageReplacement
Table = TableReplacement

__all__ = ['DocxTemplate',
           'HTML',
           'Text',
           'Table',
           'Image',
           'Document']
