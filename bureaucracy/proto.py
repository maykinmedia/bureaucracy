#!/usr/bin/env python
from collections import OrderedDict

from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER


filepath = '/home/bbt/Downloads/imtool export test 1.pptx'
save_filepath = '/home/bbt/Downloads/imtool export generated.pptx'

goat = '/home/bbt/Downloads/goat.jpg'


def construct_template_code(slide):
    """
    Build the entire template code to be rendered as one string.

    Used for debugging purposes now, to investigate the structure of
    placeholders.
    """
    template_bits = OrderedDict()

    for placeholder in slide.slide_layout.placeholders:
        template_bits[placeholder.placeholder_format.idx] = placeholder.text

    for placeholder in slide.placeholders:
        # check if there is a text already. If there is, it's an override and
        # the template bit should not be used
        if not placeholder.text:
            continue
        template_bits[placeholder.placeholder_format.idx] = placeholder.text

    print("\n".join(template_bits.values()))

    return template_bits


def main():
    pres = Presentation(filepath)

    print("Checking layouts...")
    for layout in pres.slide_layouts:
        print("  * {0.name}".format(layout))

    slides = pres.slides
    print("Presentation has {0} slide(s)".format(len(slides)))

    print("Templating out first slide...")
    slide0 = slides[0]

    print("Found placeholders:")
    for placeholder in slide0.placeholders:
        print("  * {0.placeholder_format.idx} - {0.name}".format(placeholder))

    print("Setting a test title...")
    slide0.shapes.title.text = "Example title filled in"

    construct_template_code(slide0)

    # print("Filling in an image")
    # logo_ph = slide0.placeholders[10]
    # assert logo_ph.placeholder_format.type == PP_PLACEHOLDER.PICTURE, 'Not a picture placeholder'

    pres.save(save_filepath)


if __name__ == '__main__':
    main()
