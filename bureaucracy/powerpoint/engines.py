"""
This module defines the base engine to render template fragments.
"""


class BaseEngine:

    @staticmethod
    def render(fragment, context, slide):
        raise NotImplementedError("You must implement the `render` method.")


class PythonEngine(BaseEngine):
    """
    Template engine that relies on python str.format(...).
    """

    @staticmethod
    def render(fragment, context, slide):
        return fragment.format(**context)
