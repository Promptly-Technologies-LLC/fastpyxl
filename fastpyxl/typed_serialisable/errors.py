class FieldValidationError(ValueError):
    pass


class FieldCoercionError(TypeError):
    pass


class ParseError(ValueError):
    pass


class RenderError(ValueError):
    pass
