"""
Null/Missing value validator (Issue #90)

Detects:
- NULL values (pd.isna)
- Empty strings ("")
- Whitespace-only strings ("   ")
"""

from typing import List, Optional

import pandas as pd

from .base_validator import BaseValidator, ValidationResult


class NullValidator(BaseValidator):
    """
    Validate for null/missing values in DataFrame
    """

    def __init__(self):
        super().__init__("NullValidator")

    def validate(
        self,
        df: pd.DataFrame,
        required_columns: Optional[List[str]] = None,
        check_whitespace: bool = True,
    ) -> ValidationResult:
        """
        Validate DataFrame for null/missing values

        Args:
            df: DataFrame to validate
            required_columns: List of columns that must not have nulls
            check_whitespace: Also check for whitespace-only strings

        Returns:
            ValidationResult with null value findings
        """
        issues = []
        total_cells = df.size
        null_count = 0

        # Check each column
        for col in df.columns:
            # Standard NULL check
            col_nulls = df[col].isna()
            null_indices = df[col_nulls].index.tolist()

            # Check for empty/whitespace strings
            if check_whitespace and df[col].dtype == object:
                empty_mask = df[col].apply(lambda x: isinstance(x, str) and x.strip() == "")
                empty_indices = df[empty_mask].index.tolist()
                all_null_indices = sorted(set(null_indices + empty_indices))
            else:
                all_null_indices = null_indices

            if all_null_indices:
                null_count += len(all_null_indices)

                # Check if this is a required column
                is_required = required_columns and col in required_columns

                issues.append(
                    {
                        "column": col,
                        "null_count": len(all_null_indices),
                        "null_rate": round(len(all_null_indices) / len(df), 4),
                        "row_indices": all_null_indices[:10],  # First 10 for brevity
                        "total_rows_affected": len(all_null_indices),
                        "severity": "critical" if is_required else "warning",
                    }
                )

        # Determine if passed
        if required_columns:
            # Fail if any required column has nulls
            required_issues = [i for i in issues if i["column"] in required_columns]
            passed = len(required_issues) == 0
        else:
            # Warn if more than 5% nulls
            passed = null_count / total_cells < 0.05 if total_cells > 0 else True

        # Summary
        affected_columns = len(issues)
        summary = f"Found {null_count} null values in {affected_columns} columns ({self._calculate_rate(null_count, total_cells):.1%} of total cells)"

        if required_columns:
            required_issues_count = sum(1 for i in issues if i["column"] in required_columns)
            if required_issues_count > 0:
                summary += f" - {required_issues_count} required columns have nulls!"

        return self._create_result(
            passed=passed,
            total_issues=null_count,
            total_items=total_cells,
            issues=issues,
            summary=summary,
            details={
                "total_cells": total_cells,
                "null_cells": null_count,
                "affected_columns": affected_columns,
                "required_columns": required_columns or [],
                "check_whitespace": check_whitespace,
            },
        )
