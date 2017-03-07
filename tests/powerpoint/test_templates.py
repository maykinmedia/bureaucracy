from pathlib import Path

from bureaucracy.powerpoint import Template

from .engines import ConstantEngine


TEST_FILES = Path(__file__).parent / 'files'


def test_layouts_extraction():
    test_file = str(TEST_FILES / 'template1.pptx')
    template = Template(test_file)
    assert template.layouts == ['default', 'influencers']


def test_template_render():
    test_file = str(TEST_FILES / 'template1.pptx')
    template = Template(test_file)
    template.render(context=None, render_engine=ConstantEngine)
