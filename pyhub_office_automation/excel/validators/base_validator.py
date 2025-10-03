"""
Base validator class for data validation (Issue #90)
"""

from abc import ABC, abstractmethod
from dataclasses import dataclass
from typing import Any, Dict, List, Optional

import pandas as pd


@dataclass
class ValidationResult:
    """Validation result for a single check"""

    validator_name: str
    passed: bool
    total_issues: int
    issue_rate: float  # 0.0 ~ 1.0
    issues: List[Dict[str, Any]]
    summary: str
    details: Optional[Dict[str, Any]] = None


class BaseValidator(ABC):
    """
    Abstract base class for all validators
    """

    def __init__(self, name: str):
        self.name = name

    @abstractmethod
    def validate(self, df: pd.DataFrame, **kwargs) -> ValidationResult:
        """
        Validate DataFrame and return results

        Args:
            df: DataFrame to validate
            **kwargs: Validator-specific parameters

        Returns:
            ValidationResult with validation findings
        """
        pass

    def _calculate_rate(self, issues: int, total: int) -> float:
        """Calculate issue rate (0.0 ~ 1.0)"""
        if total == 0:
            return 0.0
        return round(issues / total, 4)

    def _create_result(
        self,
        passed: bool,
        total_issues: int,
        total_items: int,
        issues: List[Dict[str, Any]],
        summary: str,
        details: Optional[Dict[str, Any]] = None,
    ) -> ValidationResult:
        """Helper to create ValidationResult"""
        return ValidationResult(
            validator_name=self.name,
            passed=passed,
            total_issues=total_issues,
            issue_rate=self._calculate_rate(total_issues, total_items),
            issues=issues,
            summary=summary,
            details=details or {},
        )
