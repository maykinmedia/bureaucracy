class SlideContainer:

    def __init__(self, slide, presentation):
        self.slide = slide
        self.presentation = presentation

    def __getattr__(self, name):
        """
        Proxy unknown attributes to the underlying slide object
        """
        return getattr(self.slide, name)

    def insert_another(self):
        """
        Inserts the same slide into the presentation after the current position.

        NOTE: there's no insert slide method, only append to the end of the
        presentation, so we're using private API here.
        """
        layout = self.slide.slide_layout
        self.presentation.slides.add_slide(layout)

        current_index = self.presentation.slides.index(self.slide)
        # always at the end, and the ID is fixed (does not indicate position)
        # move the element to behind the current slide
        slide_list = self.presentation.slides._sldIdLst
        slide_list[current_index + 1] = slide_list[-1]
