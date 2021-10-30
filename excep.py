class Error(ValueError):
    """Base class for the module"""
    pass

class InputError(Error):
    """Expression raised in violation of allowable value limit"""

    def __init__(self, message):
        self.message = message

class OptionError(Error):
    """Expression raised from selecting nonexistent option"""

    def __init__(self, message):
        self.message = message