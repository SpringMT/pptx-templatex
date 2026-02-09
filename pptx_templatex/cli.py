"""Command-line interface for pptx-templatex."""

import argparse
import sys
from pathlib import Path

from .exceptions import TemplateError
from .template_engine import TemplateEngine


def main():
    """Main entry point for the CLI."""
    parser = argparse.ArgumentParser(
        description="PowerPoint template engine with placeholder replacement",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Process with JSON config file
  pptx-templatex template.pptx config.json output.pptx

  # Use with explicit paths
  pptx-templatex /path/to/template.pptx /path/to/config.json /path/to/output.pptx

Config file format (JSON):
  {
    "slides": [
      {
        "src_page": 1,
        "replace_texts": {
          "name": "John Doe",
          "user": {"email": "john@example.com"},
          "items": [{"name": "Product A"}]
        }
      }
    ]
  }

Placeholder syntax:
  - Simple: {{ name }}
  - Nested: {{ user.name }}
  - Array: {{ items[0].name }}
"""
    )

    parser.add_argument(
        "template",
        help="Path to the template PowerPoint file (.pptx)"
    )
    parser.add_argument(
        "config",
        help="Path to the JSON configuration file"
    )
    parser.add_argument(
        "output",
        help="Path to save the output PowerPoint file (.pptx)"
    )
    parser.add_argument(
        "-v", "--version",
        action="version",
        version="%(prog)s 0.1.0"
    )

    args = parser.parse_args()

    # Validate input files exist
    template_path = Path(args.template)
    config_path = Path(args.config)

    if not template_path.exists():
        print(f"Error: Template file not found: {args.template}", file=sys.stderr)
        return 1

    if not config_path.exists():
        print(f"Error: Config file not found: {args.config}", file=sys.stderr)
        return 1

    try:
        # Process the template
        engine = TemplateEngine(template_path)
        engine.process(config_path, args.output)
        print(f"Successfully created: {args.output}")
        return 0

    except TemplateError as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1
    except Exception as e:
        print(f"Unexpected error: {str(e)}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
