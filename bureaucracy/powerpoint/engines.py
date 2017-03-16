"""
This module defines the base engine to render template fragments.
"""


class BaseEngine:

    current_slide = None
    current_placeholder = None

    def render(self, fragment, context):
        raise NotImplementedError("You must implement the `render` method.")


class PythonEngine(BaseEngine):
    """
    Template engine that relies on python str.format(...).
    """

    def render(self, fragment, context):
        return fragment.format(**context)
