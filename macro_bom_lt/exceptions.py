class IncorrectLayout(Exception):
    """exception when layout of the input excel is incorrect
    Art. /Nr. /Menge/ Beschreibung
    """

    def __init__(self, message=None):
        self.order = "1. Art. 2. Nr. 3. Menge 4. Beschreibung"
        if message is None:
            message = f" Order of columns is incorrect! Should be {self.order}"
        super().__init__(message)
