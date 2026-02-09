"""Placeholder replacement functionality with nested key support."""

import re
from typing import Any, Dict

from .exceptions import PlaceholderError


class PlaceholderReplacer:
    """Handles replacement of {{ }} placeholders with support for nested keys."""

    PLACEHOLDER_PATTERN = re.compile(r'\{\{([^}]+)\}\}')

    @staticmethod
    def get_nested_value(data: Dict[str, Any], key_path: str) -> Any:
        """
        Get value from nested dictionary using dot notation.

        Supports:
        - Simple keys: "name"
        - Nested keys: "user.name"
        - Array access: "users[0].name" or "users.[0].name"

        Args:
            data: Dictionary containing the data
            key_path: Dot-separated path to the value (e.g., "a.b[0].c")

        Returns:
            The value at the specified path

        Raises:
            PlaceholderError: If the key path is invalid or not found
        """
        # Normalize array access: convert [0] to .[0], and check for invalid indices
        # First check if there are any invalid array indices (non-numeric)
        invalid_indices = re.findall(r'\[([^\d\]]+)\]', key_path)
        if invalid_indices:
            raise PlaceholderError(
                f"Invalid array index '[{invalid_indices[0]}]' in path '{key_path}'"
            )

        normalized_path = re.sub(r'\[(\d+)\]', r'.[\1]', key_path)

        # Split by dots
        keys = normalized_path.split('.')
        current = data

        for key in keys:
            key = key.strip()
            if not key:
                continue

            # Handle array index access: [0]
            if key.startswith('[') and key.endswith(']'):
                try:
                    index = int(key[1:-1])
                    if not isinstance(current, (list, tuple)):
                        raise PlaceholderError(
                            f"Cannot index non-list value at '{key}' in path '{key_path}'"
                        )
                    if index < 0 or index >= len(current):
                        raise PlaceholderError(
                            f"Index {index} out of range for path '{key_path}'"
                        )
                    current = current[index]
                except ValueError:
                    raise PlaceholderError(
                        f"Invalid array index '{key}' in path '{key_path}'"
                    )
            else:
                # Handle dictionary key access
                if not isinstance(current, dict):
                    raise PlaceholderError(
                        f"Cannot access key '{key}' on non-dict value in path '{key_path}'"
                    )
                if key not in current:
                    raise PlaceholderError(
                        f"Key '{key}' not found in path '{key_path}'"
                    )
                current = current[key]

        return current

    @classmethod
    def replace_text(cls, text: str, replacements: Dict[str, Any]) -> str:
        """
        Replace all {{ }} placeholders in text with values from replacements dict.

        Args:
            text: Text containing placeholders
            replacements: Dictionary with replacement values

        Returns:
            Text with placeholders replaced

        Raises:
            PlaceholderError: If a placeholder cannot be replaced
        """
        def replace_match(match: re.Match) -> str:
            key_path = match.group(1).strip()
            try:
                value = cls.get_nested_value(replacements, key_path)
                return str(value)
            except PlaceholderError as e:
                raise PlaceholderError(
                    f"Failed to replace placeholder '{{{{{key_path}}}}}': {str(e)}"
                )

        return cls.PLACEHOLDER_PATTERN.sub(replace_match, text)
