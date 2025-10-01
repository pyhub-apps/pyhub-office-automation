"""
Variable management for batch scripts

Supports:
- Variable definition and retrieval
- Variable resolution in strings (${VAR} and $VAR syntax)
- Environment variable fallback
- Variable scope management
"""

import os
import re
from typing import Dict, Optional


class VariableManager:
    """Manage script variables and environment"""

    def __init__(self, initial_vars: Optional[Dict[str, str]] = None):
        """
        Initialize variable manager

        Args:
            initial_vars: Initial variable dictionary
        """
        self.variables: Dict[str, str] = initial_vars.copy() if initial_vars else {}
        self.environment = os.environ.copy()

    def set(self, name: str, value: str) -> None:
        """
        Set a variable

        Args:
            name: Variable name (alphanumeric + underscore)
            value: Variable value (string)
        """
        if not self._is_valid_name(name):
            raise ValueError(f"Invalid variable name: {name}. " "Must contain only alphanumeric characters and underscores.")
        self.variables[name] = str(value)

    def unset(self, name: str) -> None:
        """
        Remove a variable

        Args:
            name: Variable name to remove
        """
        if name in self.variables:
            del self.variables[name]

    def get(self, name: str, default: str = "") -> str:
        """
        Get variable value with fallback to environment

        Args:
            name: Variable name
            default: Default value if not found

        Returns:
            Variable value or default
        """
        # First check script variables
        if name in self.variables:
            return self.variables[name]

        # Then check environment variables
        if name in self.environment:
            return self.environment[name]

        return default

    def export(self, name: str, value: str) -> None:
        """
        Set variable and export to environment

        Args:
            name: Variable name
            value: Variable value
        """
        self.set(name, value)
        self.environment[name] = value
        os.environ[name] = value

    def resolve(self, text: str) -> str:
        """
        Resolve all variables in text

        Supports:
        - ${VAR_NAME} - Braced syntax (preferred)
        - $VAR_NAME - Unbraced syntax

        Args:
            text: Text containing variable references

        Returns:
            Text with variables resolved
        """

        def replacer(match) -> str:
            # Group 1: ${VAR} syntax
            # Group 2: $VAR syntax
            var_name = match.group(1) or match.group(2)
            return self.get(var_name)

        # Pattern: ${VAR_NAME} or $VAR_NAME
        # VAR_NAME must start with letter or underscore, followed by alphanumerics/underscores
        pattern = r"\$\{([A-Za-z_][A-Za-z0-9_]*)\}|\$([A-Za-z_][A-Za-z0-9_]*)"
        return re.sub(pattern, replacer, text)

    def has(self, name: str) -> bool:
        """
        Check if variable exists

        Args:
            name: Variable name

        Returns:
            True if variable exists in script or environment
        """
        return name in self.variables or name in self.environment

    def list_all(self) -> Dict[str, str]:
        """
        Get all variables (script variables only, not environment)

        Returns:
            Dictionary of all script variables
        """
        return self.variables.copy()

    def _is_valid_name(self, name: str) -> bool:
        """
        Validate variable name

        Rules:
        - Must start with letter or underscore
        - Can contain letters, digits, underscores
        - Cannot be empty

        Args:
            name: Variable name to validate

        Returns:
            True if valid
        """
        if not name:
            return False
        pattern = r"^[A-Za-z_][A-Za-z0-9_]*$"
        return re.match(pattern, name) is not None


def parse_set_directive(line: str) -> tuple[str, str]:
    """
    Parse @set directive

    Formats:
    - @set VAR = "value"
    - @set VAR="value"
    - @set VAR = value
    - @set VAR=value

    Args:
        line: Line containing @set directive

    Returns:
        Tuple of (variable_name, value)

    Raises:
        ValueError: If line format is invalid
    """
    # Remove @set prefix and strip
    content = line.strip().removeprefix("@set").strip()

    # Split on first = sign
    if "=" not in content:
        raise ValueError(f"Invalid @set directive: {line}. Expected format: @set VAR = value")

    name_part, value_part = content.split("=", 1)
    name = name_part.strip()
    value = value_part.strip()

    # Remove quotes if present
    if value and value[0] in ('"', "'") and value[-1] == value[0]:
        value = value[1:-1]

    if not name:
        raise ValueError(f"Invalid @set directive: {line}. Variable name cannot be empty")

    return name, value


def parse_unset_directive(line: str) -> str:
    """
    Parse @unset directive

    Format: @unset VAR_NAME

    Args:
        line: Line containing @unset directive

    Returns:
        Variable name to unset

    Raises:
        ValueError: If line format is invalid
    """
    content = line.strip().removeprefix("@unset").strip()

    if not content:
        raise ValueError(f"Invalid @unset directive: {line}. Expected format: @unset VAR_NAME")

    return content


def parse_echo_directive(line: str) -> str:
    """
    Parse @echo directive

    Format: @echo "message" or @echo message

    Args:
        line: Line containing @echo directive

    Returns:
        Message to echo (quotes removed if present)
    """
    content = line.strip().removeprefix("@echo").strip()

    # Remove quotes if present
    if content and content[0] in ('"', "'") and len(content) > 1 and content[-1] == content[0]:
        content = content[1:-1]

    return content


def parse_export_directive(line: str) -> tuple[str, str]:
    """
    Parse @export directive

    Format: @export VAR = "value"

    Args:
        line: Line containing @export directive

    Returns:
        Tuple of (variable_name, value)

    Raises:
        ValueError: If line format is invalid
    """
    # Same parsing as @set
    content = line.strip().replace("@export", "@set")
    return parse_set_directive(content)
