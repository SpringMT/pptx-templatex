"""Rich text parser for markdown-like markup in placeholder values."""

import re
from dataclasses import dataclass, field
from typing import List, Optional


@dataclass
class TextSegment:
    """A segment of text with formatting attributes."""
    text: str
    bold: bool = False
    italic: bool = False
    color: Optional[str] = None   # HEX string e.g. "#FF0000"
    size: Optional[float] = None  # font size in points


@dataclass
class RichParagraph:
    """A paragraph consisting of one or more text segments."""
    segments: List[TextSegment] = field(default_factory=list)
    bullet: bool = False


def parse_rich_text(text: str) -> List[RichParagraph]:
    """
    Parse markdown-like markup into a list of RichParagraph objects.

    Supported syntax:
        * text          -> bullet paragraph
        **text**        -> bold
        *text*          -> italic
        {color:#RRGGBB}text{/color}  -> colored text
        {size:N}text{/size}          -> font size (points)

    Args:
        text: Input string, may contain markup and newlines.

    Returns:
        List of RichParagraph objects representing the parsed content.
    """
    paragraphs: List[RichParagraph] = []
    for line in text.split("\n"):
        bullet = False
        if line.startswith("* "):
            bullet = True
            line = line[2:]
        para = RichParagraph(bullet=bullet)
        para.segments = _parse_inline(line)
        paragraphs.append(para)
    return paragraphs


def is_rich_text(text: str) -> bool:
    """Return True if the text contains any rich-text markup."""
    patterns = [
        r"\*\*.+?\*\*",
        r"(?<!\*)\*(?!\*)(?!\s).+?(?<!\s)\*(?!\*)",
        r"\{color:[^}]+\}",
        r"\{size:[^}]+\}",
        r"^[*] ",
    ]
    for pattern in patterns:
        if re.search(pattern, text, re.MULTILINE):
            return True
    return False


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

# Token types emitted by the inline tokeniser
_TOKEN_RE = re.compile(
    r"(?P<bold>\*\*)"
    r"|(?P<em>(?<!\*)\*(?!\*))"
    r"|(?P<color_open>\{color:(?P<color_val>[^}]+)\})"
    r"|(?P<color_close>\{/color\})"
    r"|(?P<size_open>\{size:(?P<size_val>[^}]+)\})"
    r"|(?P<size_close>\{/size\})"
    r"|(?P<text>[^*{]+|\{(?!\/?(?:color|size))|[*])"
)


def _parse_inline(line: str) -> List[TextSegment]:
    """Parse a single line of inline markup into TextSegment list."""
    segments: List[TextSegment] = []

    # State stack: each entry is a dict of active attributes
    bold = False
    italic = False
    color: Optional[str] = None
    size: Optional[float] = None

    pos = 0
    while pos < len(line):
        m = _TOKEN_RE.match(line, pos)
        if not m:
            pos += 1
            continue
        pos = m.end()

        if m.group("bold"):
            bold = not bold

        elif m.group("em"):
            italic = not italic

        elif m.group("color_open"):
            raw = m.group("color_val").strip()
            color = _normalise_color(raw)

        elif m.group("color_close"):
            color = None

        elif m.group("size_open"):
            try:
                size = float(m.group("size_val").strip())
            except ValueError:
                size = None

        elif m.group("size_close"):
            size = None

        elif m.group("text"):
            raw_text = m.group("text")
            if raw_text:
                segments.append(
                    TextSegment(
                        text=raw_text,
                        bold=bold,
                        italic=italic,
                        color=color,
                        size=size,
                    )
                )

    return segments if segments else [TextSegment(text=line)]


def _normalise_color(value: str) -> Optional[str]:
    """Accept '#RRGGBB' or 'RRGGBB'; return uppercase hex without '#', or None."""
    value = value.strip().lstrip("#").upper()
    if re.fullmatch(r"[0-9A-F]{6}", value):
        return value
    return None
