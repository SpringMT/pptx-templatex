"""Unit tests for PlaceholderReplacer."""

import pytest
from pptx_templatex.placeholder_replacer import PlaceholderReplacer
from pptx_templatex.exceptions import PlaceholderError


class TestGetNestedValue:
    """Tests for get_nested_value method."""

    def test_simple_key(self):
        """Test accessing a simple key."""
        data = {"name": "John", "age": 30}
        result = PlaceholderReplacer.get_nested_value(data, "name")
        assert result == "John"

    def test_nested_key_with_dot(self):
        """Test accessing nested keys with dot notation."""
        data = {"user": {"name": "John", "age": 30}}
        result = PlaceholderReplacer.get_nested_value(data, "user.name")
        assert result == "John"

    def test_deeply_nested_key(self):
        """Test accessing deeply nested keys."""
        data = {"company": {"department": {"team": {"lead": "Alice"}}}}
        result = PlaceholderReplacer.get_nested_value(data, "company.department.team.lead")
        assert result == "Alice"

    def test_array_access_with_brackets(self):
        """Test array access using bracket notation."""
        data = {"items": ["first", "second", "third"]}
        result = PlaceholderReplacer.get_nested_value(data, "items[0]")
        assert result == "first"

    def test_array_access_with_dot_brackets(self):
        """Test array access using dot-bracket notation."""
        data = {"items": ["first", "second", "third"]}
        result = PlaceholderReplacer.get_nested_value(data, "items.[1]")
        assert result == "second"

    def test_nested_array_object_access(self):
        """Test accessing nested object within array."""
        data = {"users": [{"name": "Alice", "age": 25}, {"name": "Bob", "age": 30}]}
        result = PlaceholderReplacer.get_nested_value(data, "users[0].name")
        assert result == "Alice"

    def test_complex_nested_path(self):
        """Test complex nested path with multiple arrays and objects."""
        data = {
            "company": {
                "departments": [
                    {"name": "Engineering", "teams": [{"name": "Backend"}, {"name": "Frontend"}]},
                    {"name": "Sales", "teams": [{"name": "US"}, {"name": "EU"}]}
                ]
            }
        }
        result = PlaceholderReplacer.get_nested_value(
            data, "company.departments[0].teams[1].name"
        )
        assert result == "Frontend"

    def test_key_not_found(self):
        """Test error when key doesn't exist."""
        data = {"name": "John"}
        with pytest.raises(PlaceholderError, match="Key 'age' not found"):
            PlaceholderReplacer.get_nested_value(data, "age")

    def test_nested_key_not_found(self):
        """Test error when nested key doesn't exist."""
        data = {"user": {"name": "John"}}
        with pytest.raises(PlaceholderError, match="Key 'age' not found"):
            PlaceholderReplacer.get_nested_value(data, "user.age")

    def test_array_index_out_of_range(self):
        """Test error when array index is out of range."""
        data = {"items": ["first", "second"]}
        with pytest.raises(PlaceholderError, match="Index 5 out of range"):
            PlaceholderReplacer.get_nested_value(data, "items[5]")

    def test_invalid_array_index(self):
        """Test error when array index is invalid."""
        data = {"items": ["first", "second"]}
        with pytest.raises(PlaceholderError, match="Invalid array index"):
            PlaceholderReplacer.get_nested_value(data, "items[abc]")

    def test_indexing_non_list(self):
        """Test error when trying to index non-list value."""
        data = {"name": "John"}
        with pytest.raises(PlaceholderError, match="Cannot index non-list"):
            PlaceholderReplacer.get_nested_value(data, "name[0]")

    def test_accessing_key_on_non_dict(self):
        """Test error when trying to access key on non-dict value."""
        data = {"items": ["first", "second"]}
        with pytest.raises(PlaceholderError, match="Cannot access key .* on non-dict"):
            PlaceholderReplacer.get_nested_value(data, "items.name")

    def test_array_with_tuple(self):
        """Test that tuples can be indexed like arrays."""
        data = {"items": ("first", "second", "third")}
        result = PlaceholderReplacer.get_nested_value(data, "items[1]")
        assert result == "second"

    def test_whitespace_in_key_path(self):
        """Test that whitespace in key path is handled."""
        data = {"user": {"name": "John"}}
        result = PlaceholderReplacer.get_nested_value(data, " user . name ")
        assert result == "John"


class TestReplaceText:
    """Tests for replace_text method."""

    def test_simple_replacement(self):
        """Test simple placeholder replacement."""
        text = "Hello {{ name }}!"
        replacements = {"name": "World"}
        result = PlaceholderReplacer.replace_text(text, replacements)
        assert result == "Hello World!"

    def test_multiple_replacements(self):
        """Test multiple placeholder replacements."""
        text = "{{ greeting }} {{ name }}! You are {{ age }} years old."
        replacements = {"greeting": "Hello", "name": "John", "age": 30}
        result = PlaceholderReplacer.replace_text(text, replacements)
        assert result == "Hello John! You are 30 years old."

    def test_nested_key_replacement(self):
        """Test replacement with nested keys."""
        text = "User: {{ user.name }}, Age: {{ user.age }}"
        replacements = {"user": {"name": "Alice", "age": 25}}
        result = PlaceholderReplacer.replace_text(text, replacements)
        assert result == "User: Alice, Age: 25"

    def test_array_access_replacement(self):
        """Test replacement with array access."""
        text = "First item: {{ items[0] }}, Second: {{ items[1] }}"
        replacements = {"items": ["Apple", "Banana"]}
        result = PlaceholderReplacer.replace_text(text, replacements)
        assert result == "First item: Apple, Second: Banana"

    def test_complex_nested_replacement(self):
        """Test replacement with complex nested structure."""
        text = "Department: {{ company.departments[0].name }}, Team: {{ company.departments[0].teams[0].name }}"
        replacements = {
            "company": {
                "departments": [
                    {"name": "Engineering", "teams": [{"name": "Backend"}]}
                ]
            }
        }
        result = PlaceholderReplacer.replace_text(text, replacements)
        assert result == "Department: Engineering, Team: Backend"

    def test_no_placeholders(self):
        """Test text without placeholders remains unchanged."""
        text = "This is plain text without placeholders."
        replacements = {"name": "John"}
        result = PlaceholderReplacer.replace_text(text, replacements)
        assert result == text

    def test_empty_text(self):
        """Test empty text."""
        text = ""
        replacements = {"name": "John"}
        result = PlaceholderReplacer.replace_text(text, replacements)
        assert result == ""

    def test_placeholder_not_found(self):
        """Test error when placeholder value not found."""
        text = "Hello {{ missing }}!"
        replacements = {"name": "John"}
        with pytest.raises(PlaceholderError, match="Failed to replace placeholder"):
            PlaceholderReplacer.replace_text(text, replacements)

    def test_same_placeholder_multiple_times(self):
        """Test same placeholder appearing multiple times."""
        text = "{{ name }} is {{ name }}'s name. {{ name }}!"
        replacements = {"name": "John"}
        result = PlaceholderReplacer.replace_text(text, replacements)
        assert result == "John is John's name. John!"

    def test_placeholder_with_whitespace(self):
        """Test placeholders with various whitespace."""
        text = "{{name}} {{ name }} {{  name  }}"
        replacements = {"name": "John"}
        result = PlaceholderReplacer.replace_text(text, replacements)
        assert result == "John John John"

    def test_numeric_value_conversion(self):
        """Test that numeric values are converted to strings."""
        text = "Age: {{ age }}, Score: {{ score }}"
        replacements = {"age": 30, "score": 95.5}
        result = PlaceholderReplacer.replace_text(text, replacements)
        assert result == "Age: 30, Score: 95.5"

    def test_boolean_value_conversion(self):
        """Test that boolean values are converted to strings."""
        text = "Active: {{ is_active }}"
        replacements = {"is_active": True}
        result = PlaceholderReplacer.replace_text(text, replacements)
        assert result == "Active: True"

    def test_none_value_conversion(self):
        """Test that None values are converted to strings."""
        text = "Value: {{ value }}"
        replacements = {"value": None}
        result = PlaceholderReplacer.replace_text(text, replacements)
        assert result == "Value: None"

    def test_nested_brackets_in_text(self):
        """Test that non-placeholder double brackets are preserved."""
        text = "Code: {{ code }}, Formula: {{x + y}}, Name: {{ name }}"
        replacements = {"code": "ABC", "name": "John"}
        # The middle placeholder will fail because 'x + y' is not a valid key
        with pytest.raises(PlaceholderError):
            PlaceholderReplacer.replace_text(text, replacements)

    def test_list_value_conversion(self):
        """Test that list values are converted to strings."""
        text = "Items: {{ items }}"
        replacements = {"items": [1, 2, 3]}
        result = PlaceholderReplacer.replace_text(text, replacements)
        assert result == "Items: [1, 2, 3]"

    def test_dict_value_conversion(self):
        """Test that dict values are converted to strings."""
        text = "User: {{ user }}"
        replacements = {"user": {"name": "John"}}
        result = PlaceholderReplacer.replace_text(text, replacements)
        assert "name" in result and "John" in result
