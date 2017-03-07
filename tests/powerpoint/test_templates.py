from pathlib import Path

from bureaucracy.powerpoint import Template


TEST_FILES = Path(__file__).resolve() / 'files'


def test_layouts_extraction():
    test_file = str(TEST_FILES / 'template1.pptx')
    import bpdb; bpdb.set_trace()
    template = Template(test_file)
