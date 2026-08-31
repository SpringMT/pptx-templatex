# pptx-templatex

A PowerPoint template engine for Python that provides slide copying and placeholder replacement functionality.

[日本語README](README.ja.md)

## Features

- Copy slides from a source PowerPoint file (.pptx)
- Replace placeholders enclosed in `{{ }}`
- Support for nested object access (e.g., `{{ user.name }}`, `{{ company.department.name }}`)
- Support for array element access (e.g., `{{ items[0].name }}`, `{{ users.[0].email }}`)
- Batch processing with JSON configuration files
- Table support: placeholder replacement in cells and variable-length row expansion (`replace_table_rows`)
- Image embedding: `{{ img:key }}` marker shapes replaced by images with `contain`/`cover` fitting (`replace_images`); SVG sources are embedded natively as vectors with a PNG fallback

## Installation

```bash
pip install -e .
```

For development with testing dependencies:

```bash
pip install -e ".[dev]"
```

## Quick Start

1. Create a PowerPoint template file with placeholders:
   - Open PowerPoint and create a presentation
   - Add text boxes with placeholders like `{{ name }}`, `{{ user.email }}`, `{{ items[0].title }}`
   - Save as `template.pptx`

2. Create a JSON configuration file (`config.json`):
```json
{
  "slides": [
    {
      "src_page": 1,
      "replace_texts": {
        "name": "John Doe",
        "user": {
          "email": "john@example.com"
        },
        "items": [
          {"title": "First Item"}
        ]
      }
    }
  ]
}
```

3. Run the command:
```bash
pptx-templatex template.pptx config.json output.pptx
```

## Usage

### Command Line Interface

After installation, you can use the `pptx-templatex` command:

```bash
# Basic usage
pptx-templatex template.pptx config.json output.pptx

# View help
pptx-templatex --help

# View version
pptx-templatex --version
```

### Python API

```python
from pptx_templatex import TemplateEngine

# Load template file
engine = TemplateEngine("template.pptx")

# Define configuration
config = {
    "slides": [
        {
            "src_page": 1,
            "replace_texts": {
                "name": "John Doe",
                "title": "Software Engineer"
            }
        }
    ]
}

# Process and output
engine.process(config, "output.pptx")
```

### Nested Object Access

```python
config = {
    "slides": [
        {
            "src_page": 1,
            "replace_texts": {
                "user": {
                    "name": "John Doe",
                    "email": "john@example.com"
                }
            }
        }
    ]
}
```

In the template:
```
Name: {{ user.name }}
Email: {{ user.email }}
```

### Array Element Access

```python
config = {
    "slides": [
        {
            "src_page": 1,
            "replace_texts": {
                "items": [
                    {"name": "Product A", "price": "1000"},
                    {"name": "Product B", "price": "2000"}
                ]
            }
        }
    ]
}
```

In the template:
```
First item: {{ items[0].name }} - ${{ items[0].price }}
Second item: {{ items[1].name }} - ${{ items[1].price }}
```

### Creating Multiple Slides

```python
config = {
    "slides": [
        {"src_page": 1, "replace_texts": {"title": "Introduction"}},
        {"src_page": 2, "replace_texts": {"content": "Main Content"}},
        {"src_page": 1, "replace_texts": {"title": "Conclusion"}},
    ]
}
```

### Using JSON Configuration File

Create a `config.json` file:
```json
{
  "slides": [
    {
      "src_page": 1,
      "replace_texts": {
        "name": "John Doe",
        "items": [{"key": "value"}]
      }
    }
  ]
}
```

Use with CLI:
```bash
pptx-templatex template.pptx config.json output.pptx
```

Or use with Python API:
```python
engine = TemplateEngine("template.pptx")
engine.process("config.json", "output.pptx")
```

### Real-World Example

**Scenario**: Generate personalized presentation for multiple users

1. Create `template.pptx` with:
   - Slide 1: Title slide with `{{ name }}` and `{{ title }}`
   - Slide 2: Content slide with `{{ company.name }}` and `{{ company.address }}`
   - Slide 3: List slide with `{{ items[0].description }}`

2. Create `config.json`:
```json
{
  "slides": [
    {
      "src_page": 1,
      "replace_texts": {
        "name": "Alice Smith",
        "title": "Senior Developer"
      }
    },
    {
      "src_page": 2,
      "replace_texts": {
        "company": {
          "name": "Tech Corp",
          "address": "123 Main St"
        }
      }
    },
    {
      "src_page": 3,
      "replace_texts": {
        "items": [
          {"description": "Improved performance by 50%"},
          {"description": "Reduced bugs by 30%"}
        ]
      }
    }
  ]
}
```

3. Generate:
```bash
pptx-templatex template.pptx config.json alice_presentation.pptx
```

This creates a 3-slide presentation with all placeholders replaced.

## Configuration Format

### Configuration Object Structure

```json
{
  "slides": [
    {
      "src_page": 1,
      "replace_texts": {
        "key": "value"
      }
    }
  ]
}
```

- `slides` (required): Array of slide configurations
  - `src_page` (required): Source slide number to copy (1-based index)
  - `replace_texts` (optional): Mapping of text to replace
  - `replace_table_rows` (optional): Mapping of row-set key to a list of row dicts
    for variable-length table expansion (see below)
  - `replace_images` (optional): Mapping of `{{ img:key }}` marker key to an image spec
    (see below)

### Tables

Placeholders inside native table cells are replaced the same way as text boxes
(via `replace_texts`, rich text supported).

For variable-length rows, put a single **template row** in the table whose cells
contain `{{ <key>.<field> }}` placeholders, and pass the row data via
`replace_table_rows`:

```json
{
  "slides": [
    {
      "src_page": 1,
      "replace_table_rows": {
        "rows": [
          {"c1": "Typhoon approach", "c2": "Moved schedule up", "c3": "Completed"},
          {"c1": "Speaker absent", "c2": "Switched to video", "c3": "Completed"}
        ]
      }
    }
  ]
}
```

The template row (cells containing `{{ rows.c1 }}`, `{{ rows.c2 }}`, ...) is
duplicated once per item — inheriting the row's borders, fill and fonts — and
then removed. An empty list removes the template row without inserting anything.
Rows whose key is not present in `replace_table_rows` are left untouched.

### Images

Put a text shape containing ``{{ img:key }}`` in the template (the marker). The image is
placed at the marker shape's position/size and the marker is removed:

```json
{
  "slides": [
    {
      "src_page": 1,
      "replace_images": {
        "photo1": {"path": "photos/cover.png", "fit": "cover"},
        "photo2": {"path": "photos/page.png", "fit": "contain"}
      }
    }
  ]
}
```

- Image source: ``path`` (file path) or ``data`` (bytes, programmatic use only)
- ``fit``: ``"contain"`` (default) letterboxes inside the marker bounds without cropping;
  ``"cover"`` fills the bounds with a centered crop
- Markers whose key is not present in ``replace_images`` are **removed** (no marker text
  leaks into the output). Keys are resolved per slide, so copying the same template page
  multiple times with different images just works

#### SVG images

SVG sources (a ``path`` ending in ``.svg`` or ``data`` containing an SVG document) are
embedded **natively as vectors** using the ``asvg:svgBlip`` extension, together with a
high-resolution PNG fallback rasterized at 3x the marker bounds. PowerPoint 2019 /
Microsoft 365 render the crisp vector; older viewers (and converters such as Keynote or
Google Slides) fall back to the PNG.

SVG support requires the optional ``resvg-py`` dependency:

```bash
pip install 'pptx-templatex[svg]'
```

### Placeholder Syntax

- Simple replacement: `{{ key }}`
- Nested keys: `{{ user.name }}`, `{{ company.department.name }}`
- Array access: `{{ items[0] }}`, `{{ users[0].name }}`
- Complex paths: `{{ company.departments[0].teams[1].name }}`

## Testing

```bash
pytest
```

With coverage report:

```bash
pytest --cov=pptx_templatex --cov-report=html
```

## Project Structure

```
pptx-templatex/
├── pptx_templatex/
│   ├── __init__.py
│   ├── template_engine.py       # Main template engine
│   ├── placeholder_replacer.py  # Placeholder replacement logic
│   └── exceptions.py            # Custom exceptions
├── tests/
│   ├── __init__.py
│   ├── test_template_engine.py
│   └── test_placeholder_replacer.py
├── examples/
│   ├── example_usage.py
│   └── config.json
├── pyproject.toml
└── README.md
```

## License

MIT

## Author

SpringMT
