"""PowerPoint template engine with slide copying and placeholder replacement."""

from .template_engine import TemplateEngine
from .exceptions import TemplateError, PlaceholderError

__version__ = "0.1.0"
__all__ = ["TemplateEngine", "TemplateError", "PlaceholderError"]
