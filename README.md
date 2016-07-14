
# Bureaucracy

Bureaucracy is a library that allows you to use .docx templates as
templates using MailMerge fields and can save them as docx's or pdfs. It 
can insert plain text, images, tables and (simple) HTML. See 
[django-bureaucracy](https://bitbucket.org/maykinmedia/django-bureaucracy)
for some mild django integration.

## Usage 

### Example

```python

from bureaucracy import DocxTemplate, HTML, Image, Table

doc = DocxTemplate('examples/sample.docx')

context = {
    'table': Table(data=[['this is the first cell of the first row', 'this is the second cell of the first row'],
                          ['the second row', 'etc'], 
                          ['etc', 'etc]], 
                   headers=['header 1', 'header 2']),
    'image': Image('pigeon.jpg')
    'html': HTML(<p><strong>bold</strong>-notbold</p><ul><li>hop</li><li>la</li><li>kee</li></ul>")
    'text': 'some text',
}

doc.render_and_save('generated.docx', context)
doc.render_and_save('generated.pdf', context, format='pdf')
    
```

### Inserting mail merge fields

Bureaucracy expects the .docx-files passed to the `DocxTemplate` constructor
to contain MailMerge fields whose names correspond to the ones used 
in the context dict. How this is done exactly depends on the version of
Office you have, but it seems that it's always a variation on Insert > Field > Mail Merge > Mergefield
and then entering the name:

![What it looks like on Office Mac 2015](docs/mailmerge_mac.png?raw=true "Mailmerge on mac")


## Installation


```
pip install -e git+https://bitbucket.org/maykinmedia/bureaucracy.git#egg=bureaucracy-0.1
```


Note that although this will install the pypandoc dependency, that package
makes use of the pandoc executable whose installation sometimes fails. 
To work around this, install pandoc on it's own with your favorite package manager and
make it available on the path.

For the pdf generation, bureaucracy needs the LibreOffice soffice executable 
to be installed and on the path.






