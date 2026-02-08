"""Example usage of pptx-templatex library."""

from pptx_templatex import TemplateEngine

# Example 1: Simple text replacement
def example_simple():
    """Simple placeholder replacement example."""
    print("Example 1: Simple text replacement")

    engine = TemplateEngine("template.pptx")

    config = {
        "slides": [
            {"src_page": 1, "replace_texts": {"name": "John Doe", "title": "Software Engineer"}}
        ]
    }

    engine.process(config, "output_simple.pptx")
    print("Created: output_simple.pptx\n")


# Example 2: Nested object access
def example_nested():
    """Nested key access example."""
    print("Example 2: Nested object access")

    engine = TemplateEngine("template.pptx")

    config = {
        "slides": [
            {
                "src_page": 1,
                "replace_texts": {
                    "user": {
                        "name": "Alice Smith",
                        "email": "alice@example.com",
                        "department": "Engineering"
                    },
                    "date": "2024-01-15"
                }
            }
        ]
    }

    engine.process(config, "output_nested.pptx")
    print("Created: output_nested.pptx\n")


# Example 3: Array access
def example_array():
    """Array access example."""
    print("Example 3: Array access")

    engine = TemplateEngine("template.pptx")

    config = {
        "slides": [
            {
                "src_page": 1,
                "replace_texts": {
                    "items": [
                        {"name": "Product A", "price": "$99"},
                        {"name": "Product B", "price": "$149"},
                        {"name": "Product C", "price": "$199"}
                    ]
                }
            }
        ]
    }

    engine.process(config, "output_array.pptx")
    print("Created: output_array.pptx\n")


# Example 4: Multiple slides
def example_multiple_slides():
    """Multiple slides with different replacements."""
    print("Example 4: Multiple slides")

    engine = TemplateEngine("template.pptx")

    config = {
        "slides": [
            {"src_page": 1, "replace_texts": {"title": "Introduction", "content": "Welcome"}},
            {"src_page": 2, "replace_texts": {"section": "Overview", "details": "Main points"}},
            {"src_page": 1, "replace_texts": {"title": "Conclusion", "content": "Thank you"}},
        ]
    }

    engine.process(config, "output_multiple.pptx")
    print("Created: output_multiple.pptx\n")


# Example 5: Complex nested structure
def example_complex():
    """Complex nested structure example."""
    print("Example 5: Complex nested structure")

    engine = TemplateEngine("template.pptx")

    config = {
        "slides": [
            {
                "src_page": 1,
                "replace_texts": {
                    "company": {
                        "name": "TechCorp",
                        "departments": [
                            {
                                "name": "Engineering",
                                "lead": "Alice",
                                "teams": [
                                    {"name": "Backend", "size": 10},
                                    {"name": "Frontend", "size": 8}
                                ]
                            },
                            {
                                "name": "Sales",
                                "lead": "Bob",
                                "teams": [
                                    {"name": "US", "size": 15},
                                    {"name": "EU", "size": 12}
                                ]
                            }
                        ]
                    }
                }
            }
        ]
    }

    engine.process(config, "output_complex.pptx")
    print("Created: output_complex.pptx\n")


# Example 6: Using JSON config file
def example_json_config():
    """Using JSON configuration file."""
    print("Example 6: Using JSON config file")

    engine = TemplateEngine("template.pptx")
    engine.process("config.json", "output_from_json.pptx")
    print("Created: output_from_json.pptx\n")


if __name__ == "__main__":
    print("pptx-templatex Examples")
    print("=" * 50)
    print()

    # Uncomment the examples you want to run
    # Make sure you have a template.pptx file with appropriate placeholders

    # example_simple()
    # example_nested()
    # example_array()
    # example_multiple_slides()
    # example_complex()
    # example_json_config()

    print("Done! Make sure to create a template.pptx file first.")
