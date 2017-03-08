from pathlib import Path

from pptx import Presentation

from bureaucracy.powerpoint import Template
from bureaucracy.powerpoint.engines import PythonEngine

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


def test_template_filled_in_placeholders(tmpdir):
    """
    Assert that filled in placeholders in the actual slide are not templated away.
    """
    test_file = str(TEST_FILES / 'empty-and-filled-in-placeholders.pptx')
    template = Template(test_file)
    context = {
        'language': 'Python',
    }
    template.render(context, render_engine=PythonEngine)
    outfile = str(tmpdir.join('placeholders.pptx'))
    template.save_to(outfile)

    # check that the contents are correctly templated out
    pres = Presentation(outfile)
    assert len(pres.slides) == 2

    # first slide has just plain template code in placeholders
    first_slide = pres.slides[0]
    assert len(first_slide.placeholders) == 4  # title, brand_logo, two placeholders

    ph1 = first_slide.slide_layout.placeholders[2].placeholder_format.idx
    ph2 = first_slide.slide_layout.placeholders[3].placeholder_format.idx

    # ids are fixed within a slide
    assert ph1 == 11
    assert ph2 == 12

    # compare rendered output
    assert first_slide.placeholders[ph1].text == 'A simple Python string format template'
    assert first_slide.placeholders[ph2].text == 'Another simple Python string format template'

    second_slide = pres.slides[1]
    assert first_slide.slide_layout == second_slide.slide_layout

    # compare rendered output
    assert second_slide.placeholders[ph1].text == 'Filled in placeholder – should not be replaced'
    assert second_slide.placeholders[ph2].text == 'Another simple Python string format template'
