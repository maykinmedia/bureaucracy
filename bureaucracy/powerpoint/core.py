"""
Public interface to use powerpoint presentations as export template.
"""
from collections import OrderedDict

from pptx import Presentation

from .engines import PythonEngine
from .shapes import ShapeContainer
from .slides import SlideContainer

__all__ = ['Template']


class Template:
    """
    A powerpoint presentation that serves as a template.

    :param filepath: path to the powerpoint file on disk or filelike object.
    """

    def __init__(self, pptx):
        self._presentation = Presentation(pptx)

    @staticmethod
    def extract_shapes(slide):
        """
        Extracts all the shapes from the slide and build a tree.

        Big shapes can wrap small shapes, and based on this we can re-group
        and nest them to apply an ordering to shapes (and thus placeholders).
        """
        if not slide.shapes:
            return []

        shapes = sorted(slide.shapes, key=lambda s: (s.width, s.height), reverse=True)
        wrapped_shapes = [ShapeContainer(shape) for shape in shapes]

        all_shapes = []

        while wrapped_shapes:
            shape = wrapped_shapes.pop(0)
            all_shapes.append(shape)
            for other_shape in wrapped_shapes:
                if not shape.wraps(other_shape):
                    continue
                shape.add_child(other_shape)

        # root shapes are the shapes without parent (= the biggest shapes that wrap other shapes)
        root_shapes = [shape for shape in all_shapes if shape.is_root]
        # finally, order by the center point of the shape
        root_shapes = sorted(root_shapes, key=lambda s: (s.center_y, s.center_x))
        return root_shapes

    def get_placeholder_idx_in_correct_order(self, slide, fragments):
        """
        Given a slide, determine the order of template fragment evaluation.

        The slide placeholders are grouped by shapes which indicate that a
        set of placeholders needs to be evaluated before another set. This
        nesting translates into a deterministic order - top to bottom, and
        within a horizontal row from left to right.
        """
        shapes = self.extract_shapes(slide)
        # we now have the correct order for the placeholders
        placeholders = sum((shape.get_placeholders() for shape in shapes), [])
        ordered_phs = [ph.placeholder_format.idx for ph in placeholders]
        return [idx for idx in ordered_phs if idx in fragments]

    def extract_template_code(self, slide):
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
        fragments = {}

        # set up the slide layout placeholder as template code
        for placeholder in slide.slide_layout.placeholders:
            fragments[placeholder.placeholder_format.idx] = placeholder.text

        # if a value exists for the placeholder in the slide itself, ignore the
        # template code
        for placeholder in slide.placeholders:
            if not placeholder.text:
                continue
            del fragments[placeholder.placeholder_format.idx]

        idxes = self.get_placeholder_idx_in_correct_order(slide, fragments)
        # return the template bits in the right order
        return OrderedDict((idx, fragments[idx]) for idx in idxes)

    @property
    def layouts(self):
        """
        Returns the names of the slide layouts present in the template file.
        """
        return [layout.name for layout in self._presentation.slide_layouts]

    def render(self, context, render_engine=PythonEngine):
        engine = render_engine()

        # TODO: handle repeating slides
        for slide in self._presentation.slides:
            slide = SlideContainer(slide, self._presentation)
            fragments = self.extract_template_code(slide)
            for idx, fragment in fragments.items():
                placeholder = slide.placeholders[idx]
                rendered = engine.render(fragment, context, slide)
                placeholder.text = rendered
                self._remove_empty_placeholder(slide, idx)

    @staticmethod
    def _remove_empty_placeholder(slide, idx):
        """
        If the placeholder is empty AND has zero height, remove it from the slide.
        """
        placeholder = slide.placeholders[idx]
        # only consider empty placeholders
        if placeholder.text:
            return

        # only consider placeholders with zero height
        if not placeholder.height == 0:
            return

        shape = placeholder.element
        shape.getparent().remove(shape)

    def save_to(self, outfile):
        self._presentation.save(outfile)
