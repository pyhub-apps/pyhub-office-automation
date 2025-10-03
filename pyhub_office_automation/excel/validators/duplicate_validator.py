"""
Duplicate value validator (Issue #90)

Detects:
- Duplicate rows (entire row matches)
- Duplicate values in key columns
- Duplicate patterns
"""

from typing import List, Optional

import pandas as pd

from .base_validator import BaseValidator, ValidationResult


class DuplicateValidator(BaseValidator):
    """
    Validate for duplicate values in DataFrame
    """

    def __init__(self):
        super().__init__("DuplicateValidator")

    def validate(
        self,
        df: pd.DataFrame,
        key_columns: Optional[List[str]] = None,
        check_full_rows: bool = True,
    ) -> ValidationResult:
        """
        Validate DataFrame for duplicate values

        Args:
            df: DataFrame to validate
            key_columns: Columns that should be unique (e.g., ID, email)
            check_full_rows: Also check for full row duplicates

        Returns:
            ValidationResult with duplicate findings
        """
        issues = []
        total_duplicates = 0

        # Check full row duplicates
        if check_full_rows:
            dup_mask = df.duplicated(keep=False)
            dup_indices = df[dup_mask].index.tolist()

            if dup_indices:
                dup_groups = df[dup_mask].groupby(list(df.columns)).groups
                total_duplicates += len(dup_indices)

                issues.append(
                    {
                        "type": "full_row_duplicate",
                        "duplicate_count": len(dup_indices),
                        "unique_groups": len(dup_groups),
                        "row_indices": dup_indices[:20],  # First 20
                        "severity": "warning",
                        "description": f"{len(dup_indices)} rows are complete duplicates ({len(dup_groups)} unique groups)",
                    }
                )

        # Check key column duplicates
        if key_columns:
            for col in key_columns:
                if col not in df.columns:
                    issues.append(
                        {
                            "type": "key_column_missing",
                            "column": col,
                            "severity": "error",
                            "description": f"Key column '{col}' not found in DataFrame",
                        }
                    )
                    continue

                dup_mask = df[col].duplicated(keep=False)
                dup_indices = df[dup_mask].index.tolist()

                if dup_indices:
                    dup_values = df[dup_mask][col].unique().tolist()
                    total_duplicates += len(dup_indices)

                    issues.append(
                        {
                            "type": "key_column_duplicate",
                            "column": col,
                            "duplicate_count": len(dup_indices),
                            "unique_duplicate_values": len(dup_values),
                            "row_indices": dup_indices[:20],
                            "sample_values": dup_values[:10],
                            "severity": "critical",
                            "description": f"Key column '{col}' has {len(dup_indices)} duplicate values",
                        }
                    )

        # Determine if passed
        critical_issues = [i for i in issues if i.get("severity") == "critical"]
        passed = len(critical_issues) == 0

        # Summary
        total_rows = len(df)
        if total_duplicates > 0:
            summary = f"Found {total_duplicates} duplicate entries ({self._calculate_rate(total_duplicates, total_rows):.1%})"
        else:
            summary = "No duplicates found"

        if key_columns:
            key_dup_count = sum(i.get("duplicate_count", 0) for i in issues if i.get("type") == "key_column_duplicate")
            if key_dup_count > 0:
                summary += f" - {key_dup_count} duplicates in key columns!"

        return self._create_result(
            passed=passed,
            total_issues=total_duplicates,
            total_items=total_rows,
            issues=issues,
            summary=summary,
            details={
                "total_rows": total_rows,
                "key_columns": key_columns or [],
                "check_full_rows": check_full_rows,
            },
        )
