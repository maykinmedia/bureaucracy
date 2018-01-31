import logging
import os
import warnings

from pptx.enum.shapes import PP_PLACEHOLDER

from .engines import BaseEngine

CONTEXT_KEY_FOR_PLACEHOLDER = 'PPT_CURRENT_PLACEHOLDER'

AlreadyRendered = object()

logger = logging.getLogger('bureaucracy.powerpoin')


class AlreadyRenderedException(Exception):
    pass


class PlaceholderContainer:
    def __init__(self, placeholder, fragment: str):
        self.placeholder = placeholder
        self.fragment = fragment

    def render(self, engine: BaseEngine, context: dict):
        context[CONTEXT_KEY_FOR_PLACEHOLDER] = self

        try:
            rendered = engine.render(self.fragment, context)
            if rendered is None:
                return  # TODO placeholder delete if placeholder.text is empty?

            if self.placeholder.placeholder_format.type == PP_PLACEHOLDER.PICTURE:
                rendered = rendered.strip()
                self.render_picture(rendered)
            else:
                self.placeholder.text = rendered
        except AlreadyRenderedException:
            pass

    def render_picture(self, path):
        if os.path.exists(path):
            try:
                self.placeholder = self.placeholder.insert_picture(path)
            except OSError as err:
                logger.warning("Cannot identify the image at: '{}'."
                               "Leaving it empty and passing to the next image. "
                               "Exception: cannot identify image file, errno={}".format(path, err.errno))

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

    def insert_table(self, n_rows, n_cols):
        self.placeholder = self.placeholder.insert_table(n_rows, n_cols)
        return self.placeholder

    @property
    def is_empty(self):
        if self.placeholder.has_table:
            return False

        if hasattr(self.placeholder, 'text') and self.placeholder.text:
            return False
        else:
            return self.placeholder.height == 0

    def remove(self):
        shape = self.placeholder.element
        parent = shape.getparent()
        if parent is not None:
            parent.remove(shape)
