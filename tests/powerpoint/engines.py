import re

from bureaucracy.powerpoint.engines import BaseEngine, PythonEngine


class ConstantEngine(BaseEngine):

    @staticmethod
    def render(fragment, context, slide):
        return 'Constant'


class RepeatingSlideEngine(PythonEngine):

    REPEATWHILE_TAG = re.compile(r'{% repeatwhile (?P<variable>\w+) %}')
    POP_TAG = re.compile(r'{% pop (?P<variable>\w+) as (?P<as_var>\w+) %}')

    @classmethod
    def render(cls, fragment, context, slide):
        repeat_match = cls.REPEATWHILE_TAG.match(fragment)
        if repeat_match:
            list_obj = context[repeat_match.group('variable')]
            if list_obj:  # not empty yet
                slide.insert_another()
            return ''
        pop_match = cls.POP_TAG.match(fragment)
        if pop_match:
            list_obj = context[pop_match.group('variable')]
            context[pop_match.group('as_var')] = list_obj.pop(0)
            return ''
        return super().render(fragment, context, slide)


class HyperlinkEngine(BaseEngine):

    LINK_TAG = re.compile(r'{% link (?P<link>[\w\.]+) (?P<desc>[\w\.]+) %}')

    def render(cls, fragment, context, placeholder, slide):
        match = cls.LINK_TAG.match(fragment)
        _link, _desc = match.group('link'), match.group('desc')
        link_bits = _link.split('.')
        desc_bits = _desc.split('.')
        link = getattr(context[link_bits[0]], link_bits[1])
        desc = getattr(context[desc_bits[0]], desc_bits[1])

        run = placeholder.text_frame.paragraphs[0].add_run()
        run.text = desc
        run.hyperlink.address = link
