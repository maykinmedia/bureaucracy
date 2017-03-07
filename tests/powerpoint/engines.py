from bureaucracy.powerpoint.engines import BaseEngine


class ConstantEngine(BaseEngine):

    @staticmethod
    def render(fragment, context):
        return 'Constant'
