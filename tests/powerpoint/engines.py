import re

from bureaucracy.powerpoint.engines import BaseEngine


class ConstantEngine(BaseEngine):

    @staticmethod
    def render(fragment, context):
        return 'Constant'


class RepeatingSlideEngine(BaseEngine):

    REPEATWHILE_TAG = re.compile(r'{% repeatwhile (?P<variable>\w+) %}')

    @classmethod
    def render(cls, fragment, context):
        repeat_match = cls.REPEATWHILE_TAG.match(fragment)
        if repeat_match:
            list_obj = context[repeat_match.group('variable')]

        import bpdb; bpdb.set_trace()
        return ''
