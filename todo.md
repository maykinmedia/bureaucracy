
# Document generator as reusable Django app

1. Use the library only, for manual use in views, cronjobs, etc.
    a. The library should work similar to Django's Template class.
    b. Given a Word-template and some context, a final PDF or Word document can be created.
    c. Documents should be able to render to HttpResponse directly (for download), or to file (for storing) as either Word or PDF.
    d. There should be a utility function to convert (basic) HTML to Word. Tags like: p, b/strong, i/em, br should be supported.
    e. There should be helpers to add images or tables to documents. Like: Table(data, headers) and Image(filepath) passed to the context.
    f. All unused mailmerge variables should be removed in the final document, all extra context variables (that do not have a mailmerge variable) should be ignored. Both should log a warning in those cases.
    
    Example:

        Template(template_filpath).render(type='pdf', context={
            'heading': 'The Title',
            'data1': Table(data, header),
            'graph1': Image(image_filepath),
            'conclusion': Html(obj.conclusion_text)
        })


2. Optionally add the app to INSTALLED_APPS, to provide Admin integration:
    a. There should be a setting to state the document types (just a name) with a list of mailmerge variables.
    b. In the admin, allow Word-templates to be uploaded, with at least a (setting choice) type and filepath
    c. Store meta data on the model: created, last modified, uploader
    d. When model is saved, the mailmerge variables in the settings, should be checked against the uploaded Word-template (validation).
    e. There should be an checkbox to skip validation described in (d).
    f. The model should expose functions to create a final Word document, or a final PDF document (make use of general library functions, see below).

    Example:
    
        Document.objects.get(type=type).render(type='pdf', context=context)
        

3. The gruwesome goal 1d: HTML to Word :)
    With the above goal, I can imagine the step to render a full HTML file to Word/PDF file might work different. I think this is okay, as long as the interface of the above solution works the same.
    
    We might call "pandoc" binary to convert HTML to Word. Which is probably something like a regular Django render, where the resulting HTML is passed to pandoc.
    
    Maybe the HTML helper can convert table and image tags as well?
