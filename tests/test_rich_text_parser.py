"""Unit tests for rich_text_parser."""

from pptx_templatex.rich_text_parser import is_rich_text, parse_rich_text


class TestIsRichText:
    def test_plain_text(self):
        assert is_rich_text("普通のテキスト") is False

    def test_bold(self):
        assert is_rich_text("**bold**") is True

    def test_italic(self):
        assert is_rich_text("*italic*") is True

    def test_bullet(self):
        assert is_rich_text("* item") is True

    def test_color(self):
        assert is_rich_text("{color:#FF0000}red{/color}") is True

    def test_size(self):
        assert is_rich_text("{size:24}big{/size}") is True

    def test_newline_only(self):
        assert is_rich_text("line1\nline2") is False


class TestParseRichText:
    def test_plain(self):
        paras = parse_rich_text("hello")
        assert len(paras) == 1
        assert paras[0].bullet is False
        assert paras[0].segments[0].text == "hello"

    def test_bullet_lines(self):
        paras = parse_rich_text("* item1\n* item2")
        assert len(paras) == 2
        assert paras[0].bullet is True
        assert paras[0].segments[0].text == "item1"
        assert paras[1].bullet is True

    def test_bold(self):
        paras = parse_rich_text("**bold text**")
        segs = paras[0].segments
        assert segs[0].bold is True
        assert segs[0].text == "bold text"

    def test_italic(self):
        paras = parse_rich_text("*italic text*")
        segs = paras[0].segments
        assert segs[0].italic is True
        assert segs[0].text == "italic text"

    def test_bold_and_italic_mixed(self):
        paras = parse_rich_text("**bold** and *italic*")
        segs = paras[0].segments
        assert segs[0].bold is True
        assert segs[0].text == "bold"
        assert segs[1].bold is False
        assert segs[2].italic is True
        assert segs[2].text == "italic"

    def test_color_hex(self):
        paras = parse_rich_text("{color:#FF0000}red{/color}")
        segs = paras[0].segments
        assert segs[0].color == "FF0000"
        assert segs[0].text == "red"

    def test_color_without_hash(self):
        paras = parse_rich_text("{color:3366CC}blue{/color}")
        segs = paras[0].segments
        assert segs[0].color == "3366CC"

    def test_invalid_color_returns_none(self):
        paras = parse_rich_text("{color:notacolor}text{/color}")
        segs = paras[0].segments
        assert segs[0].color is None

    def test_size(self):
        paras = parse_rich_text("{size:24}big{/size}")
        segs = paras[0].segments
        assert segs[0].size == 24.0
        assert segs[0].text == "big"

    def test_combined_bullet_bold_color(self):
        paras = parse_rich_text("* **{color:#FF0000}重要{/color}**\n* 通常")
        assert paras[0].bullet is True
        assert paras[0].segments[0].bold is True
        assert paras[0].segments[0].color == "FF0000"
        assert paras[1].bullet is True
        assert paras[1].segments[0].bold is False

    def test_multiline_mixed(self):
        text = "* item1\n* item2\n通常テキスト"
        paras = parse_rich_text(text)
        assert len(paras) == 3
        assert paras[0].bullet is True
        assert paras[1].bullet is True
        assert paras[2].bullet is False
