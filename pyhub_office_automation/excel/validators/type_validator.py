"""
Data type validator (Issue #90)

Validates:
- Expected vs actual data types
- Numeric conversion failures
- Date format validation
- Text pattern validation
"""

from datetime import datetime
from typing import Dict, List, Optional

import pandas as pd

from .base_validator import BaseValidator, ValidationResult


class TypeValidator(BaseValidator):
    """
    Validate data types in DataFrame
    """

    def __init__(self):
        super().__init__("TypeValidator")

    def validate(
        self,
        df: pd.DataFrame,
        column_types: Optional[Dict[str, str]] = None,
        strict: bool = False,
    ) -> ValidationResult:
        """
        Validate DataFrame column types

        Args:
            df: DataFrame to validate
            column_types: Expected types {column: type}
                         Types: 'int', 'float', 'str', 'date', 'bool'
            strict: Fail on any type mismatch (default: warn only)

        Returns:
            ValidationResult with type validation findings
        """
        issues = []
        total_type_errors = 0

        if not column_types:
            # Auto-detect and report current types
            for col in df.columns:
                dtype_str = str(df[col].dtype)
                issues.append(
                    {
                        "column": col,
                        "detected_type": dtype_str,
                        "severity": "info",
                        "description": f"Column '{col}' detected as {dtype_str}",
                    }
                )

            summary = f"Auto-detected types for {len(df.columns)} columns"
            passed = True
        else:
            # Validate specified types
            for col, expected_type in column_types.items():
                if col not in df.columns:
                    issues.append(
                        {
                            "column": col,
                            "expected_type": expected_type,
                            "severity": "error",
                            "description": f"Column '{col}' not found",
                        }
                    )
                    total_type_errors += 1
                    continue

                # Type validation
                validation_result = self._validate_column_type(df[col], expected_type)

                if not validation_result["valid"]:
                    total_type_errors += validation_result["error_count"]
                    issues.append(
                        {
                            "column": col,
                            "expected_type": expected_type,
                            "actual_type": str(df[col].dtype),
                            "error_count": validation_result["error_count"],
                            "error_rate": validation_result["error_rate"],
                            "sample_errors": validation_result["sample_errors"],
                            "severity": "critical" if strict else "warning",
                            "description": validation_result["description"],
                        }
                    )

            # Determine pass/fail
            critical_issues = [i for i in issues if i.get("severity") in ["critical", "error"]]
            passed = len(critical_issues) == 0

            # Summary
            if total_type_errors > 0:
                summary = f"Found {total_type_errors} type validation errors in {len(issues)} columns"
            else:
                summary = "All column types match expectations"

        return self._create_result(
            passed=passed,
            total_issues=total_type_errors,
            total_items=len(df) * len(column_types) if column_types else len(df.columns),
            issues=issues,
            summary=summary,
            details={
                "column_types": column_types or {},
                "strict_mode": strict,
            },
        )

    def _validate_column_type(self, series: pd.Series, expected_type: str) -> Dict:
        """
        Validate a single column against expected type

        Returns:
            Dict with validation results
        """
        expected_type = expected_type.lower()
        error_count = 0
        sample_errors = []

        if expected_type == "int":
            # Check if values can be converted to int
            for idx, val in series.items():
                if pd.isna(val):
                    continue
                try:
                    int(val)
                except (ValueError, TypeError):
                    error_count += 1
                    if len(sample_errors) < 5:
                        sample_errors.append({"row": int(idx), "value": str(val)[:50]})

        elif expected_type == "float":
            # Check if values can be converted to float
            for idx, val in series.items():
                if pd.isna(val):
                    continue
                try:
                    float(val)
                except (ValueError, TypeError):
                    error_count += 1
                    if len(sample_errors) < 5:
                        sample_errors.append({"row": int(idx), "value": str(val)[:50]})

        elif expected_type == "date":
            # Check if values can be parsed as dates
            for idx, val in series.items():
                if pd.isna(val):
                    continue
                try:
                    pd.to_datetime(val)
                except (ValueError, TypeError):
                    error_count += 1
                    if len(sample_errors) < 5:
                        sample_errors.append({"row": int(idx), "value": str(val)[:50]})

        elif expected_type == "str":
            # Check if all values are strings
            for idx, val in series.items():
                if pd.isna(val):
                    continue
                if not isinstance(val, str):
                    error_count += 1
                    if len(sample_errors) < 5:
                        sample_errors.append({"row": int(idx), "value": str(val)[:50], "type": type(val).__name__})

        elif expected_type == "bool":
            # Check if values are boolean
            for idx, val in series.items():
                if pd.isna(val):
                    continue
                if not isinstance(val, (bool, int)) or (isinstance(val, int) and val not in [0, 1]):
                    error_count += 1
                    if len(sample_errors) < 5:
                        sample_errors.append({"row": int(idx), "value": str(val)[:50]})

        valid = error_count == 0
        error_rate = self._calculate_rate(error_count, len(series))

        description = f"Expected {expected_type}, found {error_count} incompatible values"

        return {
            "valid": valid,
            "error_count": error_count,
            "error_rate": error_rate,
            "sample_errors": sample_errors,
            "description": description,
        }
