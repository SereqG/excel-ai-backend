class PipelineValidationError(ValueError):
    """
    Raised when pipeline schema/semantics are invalid.

    In API layer this should typically map to HTTP 400.
    """
    pass

