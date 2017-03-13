import os
import warnings

from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.shapes.placeholder import BasePlaceholder

from .engines import BaseEngine


class PlaceholderContainer:

    def __init__(self, placeholder: BasePlaceholder, fragment: str):
        self.placeholder = placeholder
        self.fragment = fragment

    def render(self, engine: BaseEngine, context: dict):
        engine.current_placeholder = self

        rendered = engine.render(self.fragment, context)
        if rendered is None:
            return  # TODO placeholder delete if placeholder.text is empty?

        if self.placeholder.placeholder_format.type == PP_PLACEHOLDER.PICTURE:
            self.render_picture(rendered)
        else:
            self.placeholder.text = rendered

    def render_picture(self, path):
        if os.path.exists(path):
            self.placeholder.insert_picture(path)
        else:
            warnings.warn("File '{}' does not exist.")

    def insert_link(self, url, description, add_break=False):
        """
        Insert a hyperlink into the placeholder.

        It is assumed that the placeholder contains only one paragraph. If more
        paragraphs exist, the link is naively inserted into the first one.

        :param add_break: True|False: whether to put the link in the same run
          or not.

        TODO: option to add link to new paragraph.
        """
        paragraph = self.placeholder.text_frame.paragraphs[0]
        run = paragraph.add_run()
        run.text = description
        run.hyperlink.address = url
        if add_break:
            run = paragraph.add_run()
            run.text = '\n'
