>>> slide0.slide_layout
<pptx.slide.SlideLayout object at 0x7f55b8c56360>
>>> layout0 = slide0.slide_layout
>>> layout0
<pptx.slide.SlideLayout object at 0x7f55b8c56360>
>>> layout0.placeholders
<pptx.shapes.shapetree.LayoutPlaceholders object at 0x7f55b6bf2d38>
>>> len(layout0.placeholders)
44
>>> layout0.placeholders[10]
<pptx.shapes.placeholder.LayoutPlaceholder object at 0x7f55b6b58438>
>>> layout0.placeholders[10].name
'Picture Placeholder 2'
>>> layout0.placeholders[10].text
'{{ card.image }}'
