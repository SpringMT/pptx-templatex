"""Main template engine for processing PowerPoint files."""

import copy
import json
import re
from pathlib import Path
from typing import Any, Dict, Union

import lxml.etree as etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from pptx.slide import Slide
from pptx.util import Pt
from pptx_slide_copier import SlideCopier

from .exceptions import TemplateError
from .placeholder_replacer import PlaceholderReplacer
from .rich_text_parser import TextSegment, is_rich_text, parse_rich_text


class TemplateEngine:
    """
    PowerPoint template engine that copies slides and replaces placeholders.

    Usage:
        engine = TemplateEngine("template.pptx")
        config = {
            "slides": [
                {"src_page": 1, "replace_texts": {"name": "John", "age": "30"}},
                {"src_page": 2, "replace_texts": {"items": [{"name": "Item 1"}]}}
            ]
        }
        engine.process(config, "output.pptx")
    """

    def __init__(self, template_path: Union[str, Path]):
        """
        Initialize the template engine with a source PowerPoint file.

        Args:
            template_path: Path to the template PPTX file

        Raises:
            TemplateError: If template file cannot be loaded
        """
        self.template_path = Path(template_path)
        if not self.template_path.exists():
            raise TemplateError(f"Template file not found: {template_path}")

        # Store the path but don't load the presentation yet
        # We'll load it fresh for each process() call to avoid state issues
        try:
            # Just validate that the file can be opened
            test_prs = Presentation(str(self.template_path))
            self.template_prs = test_prs
        except Exception as e:
            raise TemplateError(f"Failed to load template: {str(e)}")

    def _copy_slide(self, source_slide_index: int, target_prs: Presentation) -> Slide:
        """
        Copy a slide from the template to the target presentation.

        Args:
            source_slide_index: Index of the slide to copy from the template (0-based)
            target_prs: The target presentation to copy to

        Returns:
            The newly created slide in the target presentation
        """
        return SlideCopier.copy_slide(self.template_prs, source_slide_index, target_prs)

    # ------------------------------------------------------------------
    # Rich-text helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _save_run_fmt(run):
        """Return a dict of formatting attributes from a run."""
        from pptx.oxml.ns import qn
        fmt: Dict[str, Any] = {
            "name": run.font.name,
            "size": run.font.size,
            "bold": run.font.bold,
            "italic": run.font.italic,
            "underline": run.font.underline,
            "color_type": None,
            "color_rgb": None,
            "color_scheme_xml": None,
        }
        try:
            from pptx.enum.dml import MSO_COLOR_TYPE
            if hasattr(run.font.color, "type"):
                fmt["color_type"] = run.font.color.type
                if fmt["color_type"] == MSO_COLOR_TYPE.RGB:
                    fmt["color_rgb"] = run.font.color.rgb
                elif fmt["color_type"] is not None:
                    # schemeClr など RGB以外のカラー指定: XMLノードをそのまま保存
                    rPr = run._r.find(qn("a:rPr"))
                    if rPr is not None:
                        solidFill = rPr.find(qn("a:solidFill"))
                        if solidFill is not None:
                            import copy
                            fmt["color_scheme_xml"] = copy.deepcopy(solidFill)
        except Exception:
            pass
        return fmt

    @staticmethod
    def _apply_run_fmt(run, fmt: Dict[str, Any]):
        """Apply saved formatting dict to a run."""
        from pptx.oxml.ns import qn
        if fmt["name"] is not None:
            run.font.name = fmt["name"]
        if fmt["size"] is not None:
            run.font.size = fmt["size"]
        if fmt["bold"] is not None:
            run.font.bold = fmt["bold"]
        if fmt["italic"] is not None:
            run.font.italic = fmt["italic"]
        if fmt["underline"] is not None:
            run.font.underline = fmt["underline"]
        from pptx.enum.dml import MSO_COLOR_TYPE
        if fmt["color_type"] == MSO_COLOR_TYPE.RGB and fmt["color_rgb"] is not None:
            try:
                run.font.color.rgb = fmt["color_rgb"]
            except Exception:
                pass
        elif fmt["color_scheme_xml"] is not None:
            # schemeClr などをXMLレベルで復元
            try:
                import copy
                import lxml.etree as etree
                rPr = run._r.find(qn("a:rPr"))
                if rPr is None:
                    rPr = etree.SubElement(run._r, qn("a:rPr"))
                    run._r.insert(0, rPr)
                existing = rPr.find(qn("a:solidFill"))
                if existing is not None:
                    rPr.remove(existing)
                rPr.insert(0, copy.deepcopy(fmt["color_scheme_xml"]))
            except Exception:
                pass

    @staticmethod
    def _apply_segment_fmt(run, seg: TextSegment, base_fmt: Dict[str, Any]):
        """Apply base formatting then overlay segment-level markup."""
        if base_fmt["name"] is not None:
            run.font.name = base_fmt["name"]
        if base_fmt["size"] is not None:
            run.font.size = base_fmt["size"]
        if base_fmt["underline"] is not None:
            run.font.underline = base_fmt["underline"]

        # bold / italic: markup overrides base when True, else fall back to base
        run.font.bold = True if seg.bold else base_fmt["bold"]
        run.font.italic = True if seg.italic else base_fmt["italic"]

        # font size override from markup
        if seg.size is not None:
            run.font.size = Pt(seg.size)

        # color: markup takes priority, then base
        from pptx.enum.dml import MSO_COLOR_TYPE
        if seg.color is not None:
            try:
                r = int(seg.color[0:2], 16)
                g = int(seg.color[2:4], 16)
                b = int(seg.color[4:6], 16)
                run.font.color.rgb = RGBColor(r, g, b)
            except Exception:
                pass
        elif base_fmt["color_type"] == MSO_COLOR_TYPE.RGB and base_fmt["color_rgb"] is not None:
            try:
                run.font.color.rgb = base_fmt["color_rgb"]
            except Exception:
                pass
        elif base_fmt["color_scheme_xml"] is not None:
            try:
                import copy
                from pptx.oxml.ns import qn
                rPr = run._r.find(qn("a:rPr"))
                if rPr is not None:
                    existing = rPr.find(qn("a:solidFill"))
                    if existing is not None:
                        rPr.remove(existing)
                    rPr.insert(0, copy.deepcopy(base_fmt["color_scheme_xml"]))
            except Exception:
                pass

    def _expand_rich_paragraphs(self, paragraph, rich_paras: list, base_fmt: Dict[str, Any]):
        """
        Replace a single paragraph with one or more rich-text paragraphs.

        The first RichParagraph reuses the existing paragraph element;
        subsequent ones are inserted after it in the txBody.
        """
        txBody = paragraph._p.getparent()
        ref_p = paragraph._p
        insert_idx = list(txBody).index(ref_p)

        for para_idx, rich_para in enumerate(rich_paras):
            if para_idx == 0:
                # Reuse the existing paragraph element
                p_elem = ref_p
                # Clear existing runs
                paragraph.clear()
            else:
                # Create a new paragraph by deep-copying the original
                p_elem = copy.deepcopy(ref_p)
                # Remove any runs carried over from the copy
                for r in p_elem.findall(qn("a:r")):
                    p_elem.remove(r)
                insert_idx += 1
                txBody.insert(insert_idx, p_elem)

            # Apply bullet formatting
            pPr = p_elem.find(qn("a:pPr"))
            if pPr is None:
                pPr = etree.SubElement(p_elem, qn("a:pPr"))
                p_elem.insert(0, pPr)

            # Remove any existing bullet elements before (re-)applying
            for tag in (qn("a:buNone"), qn("a:buChar"), qn("a:buAutoNum")):
                for el in pPr.findall(tag):
                    pPr.remove(el)

            if rich_para.bullet:
                bu = etree.SubElement(pPr, qn("a:buChar"))
                bu.set("char", "•")
            else:
                # Explicitly suppress inherited bullets so non-bullet lines
                # inside a bullet-styled text box are rendered correctly.
                etree.SubElement(pPr, qn("a:buNone"))

            # Add a run for each segment
            for seg in rich_para.segments:
                r_elem = etree.SubElement(p_elem, qn("a:r"))
                t_elem = etree.SubElement(r_elem, qn("a:t"))
                t_elem.text = seg.text

                # Build a temporary run object to leverage python-pptx font API
                from pptx.text.text import _Run  # noqa: PLC0415
                tmp_run = _Run(r_elem, paragraph)
                self._apply_segment_fmt(tmp_run, seg, base_fmt)

    # ------------------------------------------------------------------

    def _replace_placeholders_in_slide(self, slide: Slide, replacements: Dict[str, Any]):
        """
        Replace all {{ }} placeholders in a slide's text with values.

        Since PowerPoint may split placeholders across multiple runs,
        we need to process the entire paragraph text at once.

        Args:
            slide: The slide to process
            replacements: Dictionary with replacement values
        """
        for shape in slide.shapes:
            if not hasattr(shape, "text_frame"):
                continue

            for paragraph in shape.text_frame.paragraphs:
                # Get the full paragraph text
                full_text = paragraph.text

                # Check if there are any placeholders in this paragraph
                if "{{" not in full_text or "}}" not in full_text:
                    continue

                # Replace placeholders in the full text
                try:
                    new_text = PlaceholderReplacer.replace_text(full_text, replacements)
                except Exception:
                    # If replacement fails, skip this paragraph
                    continue

                # If text hasn't changed, skip
                if new_text == full_text:
                    continue

                # Clean control characters from the replaced text
                # Convert vertical tab (0x0B) to newline - PowerPoint uses this for soft line breaks
                new_text = new_text.replace('\x0B', '\n')
                # Remove other control characters (except \n and \r)
                new_text = re.sub(r'[\x00-\x08\x0C\x0E-\x1F]', '', new_text)

                # Find the first run with defined formatting to use as reference
                reference_run = None
                for run in paragraph.runs:
                    if run.font.name is not None:
                        reference_run = run
                        break

                # If no reference found, use the first run
                if reference_run is None and len(paragraph.runs) > 0:
                    reference_run = paragraph.runs[0]

                base_fmt = self._save_run_fmt(reference_run) if reference_run else {
                    "name": None, "size": None, "bold": None, "italic": None,
                    "underline": None, "color_type": None, "color_rgb": None,
                }

                # --- Rich text path ---
                if is_rich_text(new_text):
                    rich_paras = parse_rich_text(new_text)
                    self._expand_rich_paragraphs(paragraph, rich_paras, base_fmt)
                    continue

                # --- Plain text path (original behaviour) ---
                # Clear all existing runs
                paragraph.clear()

                # Create a new run with the replaced text
                new_run = paragraph.add_run()
                new_run.text = new_text
                self._apply_run_fmt(new_run, base_fmt)

    def process(
        self,
        config: Union[Dict, str, Path],
        output_path: Union[str, Path]
    ):
        """
        Process the template with the given configuration and save to output file.

        Args:
            config: Configuration dict or path to JSON config file
            output_path: Path to save the output PPTX file

        Raises:
            TemplateError: If processing fails
        """
        # Load config if it's a file path
        if isinstance(config, (str, Path)):
            config_path = Path(config)
            if not config_path.exists():
                raise TemplateError(f"Config file not found: {config}")
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
            except Exception as e:
                raise TemplateError(f"Failed to load config: {str(e)}")

        # Validate config
        if not isinstance(config, dict) or "slides" not in config:
            raise TemplateError("Config must contain 'slides' key")

        slides_config = config["slides"]
        if not isinstance(slides_config, list):
            raise TemplateError("'slides' must be a list")

        # Create new presentation based on template to preserve theme and layouts
        # This ensures theme fonts, colors, layouts, and other settings are maintained
        output_prs = Presentation(str(self.template_path))

        # Remove all slides from the template
        while len(output_prs.slides) > 0:
            rId = output_prs.slides._sldIdLst[0].rId
            output_prs.part.drop_rel(rId)
            del output_prs.slides._sldIdLst[0]

        # Process each slide configuration
        for idx, slide_config in enumerate(slides_config):
            if not isinstance(slide_config, dict):
                raise TemplateError(f"Slide config at index {idx} must be a dict")

            if "src_page" not in slide_config:
                raise TemplateError(f"Slide config at index {idx} missing 'src_page'")

            src_page = slide_config["src_page"]
            replace_texts = slide_config.get("replace_texts", {})

            # Validate src_page
            if not isinstance(src_page, int) or src_page < 1:
                raise TemplateError(
                    f"Invalid src_page {src_page} at index {idx}: must be positive integer"
                )

            if src_page > len(self.template_prs.slides):
                raise TemplateError(
                    f"src_page {src_page} at index {idx} exceeds template slides count "
                    f"({len(self.template_prs.slides)})"
                )

            # Copy the slide from template_prs to output_prs
            # Since output_prs is also based on the same template, layouts should match
            new_slide = self._copy_slide(src_page - 1, output_prs)

            # Note: Font normalization is no longer needed because the theme is preserved
            # from the template. Font name: None will use the theme's default font.
            # self._normalize_fonts_in_slide(new_slide)

            # Replace placeholders
            if replace_texts:
                self._replace_placeholders_in_slide(new_slide, replace_texts)

        # Save output
        try:
            output_prs.save(str(output_path))
        except Exception as e:
            raise TemplateError(f"Failed to save output: {str(e)}")
