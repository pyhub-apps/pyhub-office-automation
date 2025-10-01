"""
Control flow management for batch scripts

Supports:
- Conditional execution (@if/@elif/@else/@endif)
- Loop execution (@foreach/@endforeach, @while/@endwhile)
- Condition evaluation (file existence, comparisons, boolean logic)
"""

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, List, Optional

from .variables import VariableManager


@dataclass
class ControlBlock:
    """Represents a control flow block"""

    block_type: str  # "if", "foreach", "while"
    start_line: int
    end_line: int
    condition: Optional[str] = None
    iterator_var: Optional[str] = None  # For foreach
    iterator_list: Optional[List[str]] = None  # For foreach
    children: List["ControlBlock"] = field(default_factory=list)
    elif_blocks: List[tuple[int, str]] = field(default_factory=list)  # For if: (line_num, condition)
    else_line: Optional[int] = None  # For if


class ConditionEvaluator:
    """Evaluate conditions in control flow statements"""

    def __init__(self, var_manager: VariableManager):
        self.var_manager = var_manager

    def evaluate(self, condition: str) -> bool:
        """
        Evaluate a condition string and return boolean result

        Supports:
        - exists("path") - File existence check
        - VAR == "value" - Equality comparison
        - VAR != "value" - Inequality comparison
        - VAR > 10 - Numeric comparison
        - VAR < 10 - Numeric comparison
        - VAR >= 10 - Greater or equal
        - VAR <= 10 - Less or equal
        - condition and condition - Logical AND
        - condition or condition - Logical OR
        - not condition - Logical NOT
        - true/false - Boolean literals

        Args:
            condition: Condition string to evaluate

        Returns:
            Boolean result
        """
        # Resolve variables first
        resolved = self.var_manager.resolve(condition.strip())

        # Handle boolean literals
        if resolved.lower() in ("true", "1", "yes"):
            return True
        if resolved.lower() in ("false", "0", "no", ""):
            return False

        # Handle exists() function
        if "exists(" in resolved:
            return self._evaluate_exists(resolved)

        # Handle logical operators (and, or, not)
        if " and " in resolved:
            return self._evaluate_and(resolved)
        if " or " in resolved:
            return self._evaluate_or(resolved)
        if resolved.startswith("not "):
            return not self.evaluate(resolved[4:].strip())

        # Handle comparison operators
        for op in ("==", "!=", ">=", "<=", ">", "<"):
            if op in resolved:
                return self._evaluate_comparison(resolved, op)

        # If no operator found, treat as variable existence check
        # Non-empty string = true, empty = false
        return bool(resolved and resolved.strip())

    def _evaluate_exists(self, condition: str) -> bool:
        """Evaluate exists("path") function"""
        match = re.search(r'exists\s*\(\s*["\'](.+?)["\']\s*\)', condition)
        if not match:
            raise ValueError(f"Invalid exists() syntax: {condition}")

        file_path = match.group(1)
        # Resolve variables in path
        resolved_path = self.var_manager.resolve(file_path)
        return Path(resolved_path).exists()

    def _evaluate_and(self, condition: str) -> bool:
        """Evaluate logical AND"""
        parts = condition.split(" and ")
        return all(self.evaluate(part.strip()) for part in parts)

    def _evaluate_or(self, condition: str) -> bool:
        """Evaluate logical OR"""
        parts = condition.split(" or ")
        return any(self.evaluate(part.strip()) for part in parts)

    def _evaluate_comparison(self, condition: str, operator: str) -> bool:
        """Evaluate comparison operators"""
        parts = condition.split(operator, 1)
        if len(parts) != 2:
            raise ValueError(f"Invalid comparison: {condition}")

        left = parts[0].strip()
        right = parts[1].strip()

        # Remove quotes from string literals
        if right.startswith('"') and right.endswith('"'):
            right = right[1:-1]
        elif right.startswith("'") and right.endswith("'"):
            right = right[1:-1]

        # Try numeric comparison first
        try:
            left_num = float(left)
            right_num = float(right)
            return self._compare_numeric(left_num, right_num, operator)
        except ValueError:
            # Fall back to string comparison
            return self._compare_string(left, right, operator)

    def _compare_numeric(self, left: float, right: float, operator: str) -> bool:
        """Compare numeric values"""
        if operator == "==":
            return left == right
        elif operator == "!=":
            return left != right
        elif operator == ">":
            return left > right
        elif operator == "<":
            return left < right
        elif operator == ">=":
            return left >= right
        elif operator == "<=":
            return left <= right
        else:
            raise ValueError(f"Unknown operator: {operator}")

    def _compare_string(self, left: str, right: str, operator: str) -> bool:
        """Compare string values"""
        if operator == "==":
            return left == right
        elif operator == "!=":
            return left != right
        else:
            raise ValueError(f"String comparison only supports == and !=, got: {operator}")


def parse_if_condition(line: str) -> str:
    """
    Parse @if directive and extract condition

    Format: @if condition

    Args:
        line: Line containing @if directive

    Returns:
        Condition string

    Raises:
        ValueError: If line format is invalid
    """
    content = line.strip().removeprefix("@if").strip()
    if not content:
        raise ValueError(f"Invalid @if directive: {line}. Expected format: @if condition")
    return content


def parse_elif_condition(line: str) -> str:
    """
    Parse @elif directive and extract condition

    Format: @elif condition

    Args:
        line: Line containing @elif directive

    Returns:
        Condition string

    Raises:
        ValueError: If line format is invalid
    """
    content = line.strip().removeprefix("@elif").strip()
    if not content:
        raise ValueError(f"Invalid @elif directive: {line}. Expected format: @elif condition")
    return content


def parse_foreach_loop(line: str) -> tuple[str, str]:
    """
    Parse @foreach directive and extract iterator variable and list

    Format: @foreach var in list
            @foreach var in ${LIST_VAR}
            @foreach var in ["item1", "item2", "item3"]

    Args:
        line: Line containing @foreach directive

    Returns:
        Tuple of (iterator_var, list_expression)

    Raises:
        ValueError: If line format is invalid
    """
    content = line.strip().removeprefix("@foreach").strip()

    # Match: var in expression
    match = re.match(r"(\w+)\s+in\s+(.+)", content)
    if not match:
        raise ValueError(f"Invalid @foreach directive: {line}. Expected format: @foreach var in list")

    var_name = match.group(1)
    list_expr = match.group(2).strip()

    return var_name, list_expr


def parse_while_condition(line: str) -> str:
    """
    Parse @while directive and extract condition

    Format: @while condition

    Args:
        line: Line containing @while directive

    Returns:
        Condition string

    Raises:
        ValueError: If line format is invalid
    """
    content = line.strip().removeprefix("@while").strip()
    if not content:
        raise ValueError(f"Invalid @while directive: {line}. Expected format: @while condition")
    return content


def parse_list_expression(list_expr: str, var_manager: VariableManager) -> List[str]:
    """
    Parse list expression and return list of strings

    Supports:
    - JSON array: ["item1", "item2", "item3"]
    - Variable reference: ${LIST_VAR}
    - Space-separated: item1 item2 item3

    Args:
        list_expr: List expression string
        var_manager: Variable manager for resolving variables

    Returns:
        List of string items
    """
    # Resolve variables first
    resolved = var_manager.resolve(list_expr)

    # JSON array format
    if resolved.startswith("[") and resolved.endswith("]"):
        import json

        try:
            items = json.loads(resolved)
            return [str(item) for item in items]
        except json.JSONDecodeError:
            raise ValueError(f"Invalid JSON array: {resolved}")

    # Comma-separated
    if "," in resolved:
        return [item.strip().strip('"').strip("'") for item in resolved.split(",")]

    # Space-separated (fallback)
    return resolved.split()


def parse_onerror_directive(line: str) -> str:
    """
    Parse @onerror directive and extract error mode

    Format: @onerror continue|abort

    Args:
        line: Line containing @onerror directive

    Returns:
        Error mode string ("continue" or "abort")

    Raises:
        ValueError: If line format is invalid
    """
    content = line.strip().removeprefix("@onerror").strip()

    if content not in ("continue", "abort"):
        raise ValueError(f"Invalid @onerror directive: {line}. Expected: @onerror continue|abort")

    return content
