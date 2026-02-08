"""Custom exceptions for pptx-templatex."""


class TemplateError(Exception):
    """Base exception for template processing errors."""
    pass


class PlaceholderError(TemplateError):
    """Exception raised when placeholder replacement fails."""
    pass
