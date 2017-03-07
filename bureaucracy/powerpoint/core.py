"""
Public interface to use powerpoint presentations as export template.
"""
from collections import OrderedDict

from pptx import Presentation

__all__ = ['Template']


class Template:
    """
    A powerpoint presentation that serves as a template.

    :param filepath: path to the powerpoint file on disk or filelike object.
    """

    def __init__(self, pptx):
        self._presentation = Presentation(pptx)

    @staticmethod
    def extract_template_code(slide):
        """
        Extract the template code from slide placeholders.

        Placeholders have multiple 'levels': a placeholder can exist on a
        slide layout and on the slide itself. If the placeholder is filled in
        on the slide itself, it is not considered to be template code. If the
        value on the slide itself is empty, the value from the layout is taken
        and assumed to be template code.

        :return: an OrderedDict with placeholder id's as key and template code
          as value.
        """
        fragments = OrderedDict()

        # set up the slide layout placeholder as template code
        for placeholder in slide.slide_layout.placeholders:
            fragments[placeholder.placeholder_format.idx] = placeholder.text

        # if a value exists for the placeholder in the slide itself, overwrite
        # the template code
        for placeholder in slide.placeholders:
            if not placeholder.text:
                continue
            fragments[placeholder.placeholder_format.idx] = placeholder.text

        return fragments

    @property
    def layouts(self):
        """
        Returns the names of the slide layouts present in the template file.
        """
        return [layout.name for layout in self._presentation.slide_layouts]
