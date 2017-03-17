from tests.powerpoint.test_templates import TEST_FILES

from bureaucracy.powerpoint import Template
from bureaucracy.powerpoint.placeholders import PlaceholderContainer
from bureaucracy.powerpoint.slides import SlideContainer


def test_insert_another_slide():
    test_file = str(TEST_FILES / 'template1.pptx')
    template = Template(test_file)

    assert len(template._presentation.slides) == 1

    slide = SlideContainer(template._presentation.slides[0], template._presentation)
    slide.insert_another()

    assert len(template._presentation.slides) == 2

    assert len(template._presentation.slides[0].shapes) == len(template._presentation.slides[1].shapes)

def test_insert_link():
    test_file = str(TEST_FILES / 'hyperlink.pptx')
    template = Template(test_file)

    ph = list(template._presentation.slides[0].placeholders)[0]
    shape = PlaceholderContainer(ph, "")

    shape.insert_link("http://www.whygodwhy.com", "foo")

    assert ph.text == "foo"
    assert ph.text_frame.paragraphs[0].runs[0].hyperlink.address == "http://www.whygodwhy.com"
