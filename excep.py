class Error(ValueError):
    """Base class for the module"""
    pass

class InputError(Error):
    """Expression raised for errors in input"""

    def __init__(self, message):
        self.message = message
