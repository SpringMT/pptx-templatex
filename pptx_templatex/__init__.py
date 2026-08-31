"""PowerPoint template engine with slide copying and placeholder replacement."""

from .exceptions import PlaceholderError, TemplateError
from .template_engine import TemplateEngine

__version__ = "0.5.0"
__all__ = ["TemplateEngine", "TemplateError", "PlaceholderError"]
