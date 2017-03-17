class ShapeContainer:
    def __init__(self, shape):
        self.shape = shape
        self.children = []
        self.parents = []

    def wraps(self, other_shape):
        # other x coordinates must be less or equal than this ones
        if other_shape.x1 < self.x1 or other_shape.x2 > self.x2:
            return False
        # other y coordinates must be less or equal than this ones
        if other_shape.y1 < self.y1 or other_shape.y2 > self.y2:
            return False
        return True

    def add_child(self, other_shape):
        self.children.append(other_shape)
        other_shape.parents.append(self)

    @property
    def x1(self):
        return self.shape.left

    @property
    def x2(self):
        return self.shape.left + self.shape.width

    @property
    def y1(self):
        return self.shape.top

    @property
    def y2(self):
        return self.shape.top + self.shape.height

    @property
    def center_x(self):
        return self.shape.left + self.shape.width / 2

    @property
    def center_y(self):
        return self.shape.top + self.shape.height / 2

    @property
    def is_root(self):
        return not self.parents

    def get_placeholders(self):
        """
        Flatten the nested structure and return the placeholders in the correct order.
        """
        placeholders = [self.shape] if self.shape.is_placeholder else []
        children = sorted(self.children, key=lambda s: (s.center_y, s.center_x))
        for child in children:
            placeholders += child.get_placeholders()
        return placeholders
