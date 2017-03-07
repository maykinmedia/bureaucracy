from pathlib import Path

from pptx import Presentation

from bureaucracy.powerpoint import Template

from .engines import ConstantEngine


TEST_FILES = Path(__file__).parent / 'files'


def test_layouts_extraction():
    test_file = str(TEST_FILES / 'template1.pptx')
    template = Template(test_file)
    assert template.layouts == ['default', 'influencers']


def test_template_render(tmpdir):
    test_file = str(TEST_FILES / 'template1.pptx')
    template = Template(test_file)
    template.render(context=None, render_engine=ConstantEngine)

    outfile = str(tmpdir.join('constant-engine.pptx'))
    template.save_to(outfile)

    # check that the contents are correctly templated out
    pres = Presentation(outfile)
    assert len(pres.slides) == 1
    slide = pres.slides[0]
    assert slide.slide_layout.name == 'influencers'

    for ph in slide.placeholders:
        assert ph.text == 'Constant'
