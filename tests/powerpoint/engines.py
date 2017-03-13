import re

from bureaucracy.powerpoint.engines import BaseEngine, PythonEngine


class ConstantEngine(BaseEngine):

    @staticmethod
    def render(fragment, context):
        return 'Constant'


class RepeatingSlideEngine(PythonEngine):

    REPEATWHILE_TAG = re.compile(r'{% repeatwhile (?P<variable>\w+) %}')
    POP_TAG = re.compile(r'{% pop (?P<variable>\w+) as (?P<as_var>\w+) %}')

    def render(self, fragment, context):
        repeat_match = self.REPEATWHILE_TAG.match(fragment)
        if repeat_match:
            list_obj = context[repeat_match.group('variable')]
            if list_obj:  # not empty yet
                self.current_slide.insert_another()
            return ''
        pop_match = self.POP_TAG.match(fragment)
        if pop_match:
            list_obj = context[pop_match.group('variable')]
            context[pop_match.group('as_var')] = list_obj.pop(0)
            return ''
        return super().render(fragment, context)


class HyperlinkEngine(BaseEngine):

    LINK_TAG = re.compile(r'{% link (?P<link>[\w\.]+) (?P<desc>[\w\.]+) %}')

    def render(self, fragment, context):
        match = self.LINK_TAG.match(fragment)
        _link, _desc = match.group('link'), match.group('desc')
        link_bits = _link.split('.')
        desc_bits = _desc.split('.')
        link = getattr(context[link_bits[0]], link_bits[1])
        desc = getattr(context[desc_bits[0]], desc_bits[1])
        self.current_placeholder.insert_link(link, desc)


class HyperlinkEngine2(BaseEngine):

    LINK_TAG = re.compile(r'{% link (?P<link>[\w\.]+) (?P<desc>[\w\.]+)( add_break=(?P<addbreak>(True|False)))? %}')
    FOR_TAG = re.compile(r'{% for (?P<local>\w+) in (?P<container>\w+) %}'
                         r'(?P<subtag>.*)'
                         r'{% endfor %}', re.DOTALL)

    def render(self, fragment, context):
        match = self.FOR_TAG.match(fragment)

        local_varname = match.group('local')
        subtag = match.group('subtag')
        container = context[match.group('container')]

        match2 = self.LINK_TAG.search(subtag)
        _link, _desc = match2.group('link'), match2.group('desc')
        add_break = {'True': True, 'False': False}[match2.group('addbreak')]
        link_bits = _link.split('.')
        desc_bits = _desc.split('.')

        for item in container:
            context[local_varname] = item
            link = getattr(context[link_bits[0]], link_bits[1])
            desc = getattr(context[desc_bits[0]], desc_bits[1])
            self.current_placeholder.insert_link(link, desc, add_break=add_break)
