#!/bin/bash
# Script to test pptx-templatex installation in an isolated environment

set -e

echo "====== Testing pptx-templatex installation ======"
echo ""

# Create temporary directory for testing
TEST_DIR=$(mktemp -d)
echo "Creating test directory: $TEST_DIR"
cd "$TEST_DIR"

# Create virtual environment
echo "Creating virtual environment..."
python3 -m venv venv
source venv/bin/activate

# Get the directory where this script is located
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"

# Install the package
echo ""
echo "Installing pptx-templatex from local path..."
pip install -q "$SCRIPT_DIR"

# Verify installation
echo ""
echo "Verifying installation..."
which pptx-templatex
pptx-templatex --version

# Create test template
echo ""
echo "Creating test template..."
python3 << 'EOF'
from pptx import Presentation
from pptx.util import Inches

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])
textbox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(1))
textbox.text_frame.text = "Hello {{ name }}! Welcome to {{ company.name }}."

textbox2 = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(5), Inches(1))
textbox2.text_frame.text = "First item: {{ items[0].title }}"

prs.save("template.pptx")
print("✓ Created template.pptx")
EOF

# Create config file
echo ""
echo "Creating config file..."
cat > config.json << 'EOF'
{
  "slides": [
    {
      "src_page": 1,
      "replace_texts": {
        "name": "Alice",
        "company": {
          "name": "TechCorp"
        },
        "items": [
          {"title": "Innovation"}
        ]
      }
    }
  ]
}
EOF
echo "✓ Created config.json"

# Run the command
echo ""
echo "Running pptx-templatex command..."
pptx-templatex template.pptx config.json output.pptx
echo "✓ Created output.pptx"

# Verify output
echo ""
echo "Verifying output..."
python3 << 'EOF'
from pptx import Presentation

prs = Presentation("output.pptx")
texts = []
for slide in prs.slides:
    for shape in slide.shapes:
        if hasattr(shape, "text"):
            texts.append(shape.text)

text_str = " ".join(texts)
print(f"Output contains: {text_str}")

assert "Alice" in text_str, "Name replacement failed"
assert "TechCorp" in text_str, "Company name replacement failed"
assert "Innovation" in text_str, "Array item replacement failed"
assert "{{ name }}" not in text_str, "Placeholder not replaced"

print("✓ All replacements successful!")
EOF

# Test Python API
echo ""
echo "Testing Python API..."
python3 << 'EOF'
from pptx_templatex import TemplateEngine

engine = TemplateEngine("template.pptx")
config = {
    "slides": [
        {
            "src_page": 1,
            "replace_texts": {
                "name": "Bob",
                "company": {"name": "StartupXYZ"},
                "items": [{"title": "Growth"}]
            }
        }
    ]
}
engine.process(config, "output2.pptx")
print("✓ Python API test successful!")
EOF

# Cleanup
echo ""
echo "Cleaning up..."
deactivate
cd -
rm -rf "$TEST_DIR"

echo ""
echo "====== All tests passed! ======"
echo "Installation directory was: $TEST_DIR (now removed)"
