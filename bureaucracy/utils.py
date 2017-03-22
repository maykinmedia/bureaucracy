from lxml.etree import tostring


NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
}


def print_node(node):
    print(tostring(node, pretty_print=True).decode('utf-8'))


def namespaced(name, ns='w'):
    return '{{{0}}}{1}'.format(NAMESPACES[ns], name)


DOCX_MIMETYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
PDF_MIMETYPE = 'application/pdf'
PPTX_MIMETYPE = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'

DOCUMENT_MIME_TYPES = {
    'docx': DOCX_MIMETYPE,
    'pdf': PDF_MIMETYPE,
    'pptx': PPTX_MIMETYPE
}