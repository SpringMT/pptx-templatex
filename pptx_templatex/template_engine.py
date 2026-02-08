"""Main template engine for processing PowerPoint files."""

import json
import re
from pathlib import Path
from typing import Any, Dict, List, Union
from pptx import Presentation
from pptx.util import Inches
from pptx.slide import Slide
from .placeholder_replacer import PlaceholderReplacer
from .slide_copier import SlideCopier
from .exceptions import TemplateError


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

    def _normalize_fonts_in_slide(self, slide: Slide):
        """
        Normalize font properties in a slide, fixing runs with None font names.

        When PowerPoint creates runs for special characters like {{ }}, it may
        leave font.name as None. This method inherits font properties from the
        first run in the paragraph that has a defined font, or applies a default.

        Args:
            slide: The slide to process
        """
        # First, collect all defined fonts in the slide to use as reference
        slide_fonts = set()
        for shape in slide.shapes:
            if hasattr(shape, "text_frame"):
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if run.text.strip() and run.font.name is not None:
                            slide_fonts.add(run.font.name)

        # Determine the most common font in the slide (or use first one found)
        default_font = slide_fonts.pop() if slide_fonts else "Meiryo UI"

        for shape in slide.shapes:
            if not hasattr(shape, "text_frame"):
                continue

            for paragraph in shape.text_frame.paragraphs:
                # Skip empty paragraphs
                if not paragraph.text.strip():
                    continue

                # Find the first run with a defined font name to use as reference
                reference_font_name = None
                for run in paragraph.runs:
                    if run.text.strip() and run.font.name is not None:
                        reference_font_name = run.font.name
                        break

                # If no reference font found in this paragraph, use the slide's default
                if reference_font_name is None:
                    reference_font_name = default_font

                # Apply the reference font to all runs with None font name
                for run in paragraph.runs:
                    if run.font.name is None:
                        run.font.name = reference_font_name

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
                except:
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

                # Save formatting from reference run
                if reference_run:
                    ref_font_name = reference_run.font.name
                    ref_font_size = reference_run.font.size
                    ref_font_bold = reference_run.font.bold
                    ref_font_italic = reference_run.font.italic
                    ref_font_underline = reference_run.font.underline

                    # Save color
                    ref_color_type = None
                    ref_color_rgb = None
                    try:
                        if hasattr(reference_run.font.color, 'type'):
                            ref_color_type = reference_run.font.color.type
                            if ref_color_type == 1:  # RGB
                                ref_color_rgb = reference_run.font.color.rgb
                    except:
                        pass
                else:
                    ref_font_name = "Arial"
                    ref_font_size = None
                    ref_font_bold = None
                    ref_font_italic = None
                    ref_font_underline = None
                    ref_color_type = None
                    ref_color_rgb = None

                # Clear all existing runs
                paragraph.clear()

                # Create a new run with the replaced text
                new_run = paragraph.add_run()
                new_run.text = new_text

                # Apply the saved formatting
                if ref_font_name is not None:
                    new_run.font.name = ref_font_name
                if ref_font_size is not None:
                    new_run.font.size = ref_font_size
                if ref_font_bold is not None:
                    new_run.font.bold = ref_font_bold
                if ref_font_italic is not None:
                    new_run.font.italic = ref_font_italic
                if ref_font_underline is not None:
                    new_run.font.underline = ref_font_underline

                # Apply color
                if ref_color_type == 1 and ref_color_rgb is not None:
                    try:
                        new_run.font.color.rgb = ref_color_rgb
                    except:
                        pass

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

            # Normalize font properties (fix None font names) - first pass
            self._normalize_fonts_in_slide(new_slide)

            # Replace placeholders
            if replace_texts:
                self._replace_placeholders_in_slide(new_slide, replace_texts)

            # Normalize font properties again after replacement
            # This ensures all text (including non-placeholder text) has proper fonts
            self._normalize_fonts_in_slide(new_slide)

        # Save output
        try:
            output_prs.save(str(output_path))
        except Exception as e:
            raise TemplateError(f"Failed to save output: {str(e)}")
