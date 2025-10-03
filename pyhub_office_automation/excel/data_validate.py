"""
Data validation CLI command (Issue #90)

Comprehensive data quality validation for Excel ranges.
"""

import json
from enum import Enum
from typing import List, Optional

import pandas as pd
import typer
from rich.console import Console
from rich.table import Table

from pyhub_office_automation.excel.utils import get_or_open_workbook, get_sheet
from pyhub_office_automation.version import get_version

from .validators import DuplicateValidator, NullValidator, TypeValidator

console = Console()


class OutputFormat(str, Enum):
    """Output format options"""

    JSON = "json"
    TEXT = "text"


class ValidationCheck(str, Enum):
    """Available validation checks"""

    NULL = "null"
    DUPLICATE = "duplicate"
    TYPE = "type"
    ALL = "all"


def data_validate(
    range_addr: str = typer.Option(..., "--range", help="Range to validate (e.g., A1:Z1000, table name)"),
    sheet: Optional[str] = typer.Option(None, "--sheet", help="Sheet name"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help="Workbook name (if multiple open)"),
    file_path: Optional[str] = typer.Option(None, "--file-path", help="Excel file path"),
    checks: str = typer.Option("all", "--checks", help="Validation checks to run (comma-separated: null,duplicate,type,all)"),
    required_columns: Optional[str] = typer.Option(
        None, "--required-columns", help="Columns that must not be null (comma-separated)"
    ),
    key_columns: Optional[str] = typer.Option(None, "--key-columns", help="Columns that must be unique (comma-separated)"),
    column_types: Optional[str] = typer.Option(
        None, "--column-types", help='Expected types (format: "col1:int,col2:date,col3:str")'
    ),
    strict_types: bool = typer.Option(False, "--strict-types", help="Fail on any type mismatch"),
    output_format: OutputFormat = typer.Option(OutputFormat.TEXT, "--format", help="Output format (json/text)"),
):
    """
    Validate Excel data quality

    Performs comprehensive validation checks on Excel data:
    - Null/missing value detection
    - Duplicate row/column detection
    - Data type validation

    \\b
    Examples:
      # Basic validation (all checks)
      oa excel data-validate --range "A1:Z1000"

      # Specific checks only
      oa excel data-validate --range "A1:Z1000" --checks null,duplicate

      # Required columns validation
      oa excel data-validate --range "A1:Z100" --required-columns "이름,이메일,전화번호"

      # Key column uniqueness
      oa excel data-validate --range "A1:Z100" --key-columns "회원ID,이메일"

      # Type validation
      oa excel data-validate --range "A1:Z100" --column-types "나이:int,가격:float,날짜:date"

      # JSON output for AI agents
      oa excel data-validate --range "A1:Z100" --format json
    """
    try:
        # Get workbook and sheet
        book = get_or_open_workbook(file_path=file_path, workbook_name=workbook_name)
        sht = get_sheet(book, sheet)

        # Read data
        data_range = sht.range(range_addr)
        values = data_range.value

        # Convert to DataFrame
        if not values:
            raise ValueError(f"No data found in range {range_addr}")

        if isinstance(values[0], list):
            df = pd.DataFrame(values[1:], columns=values[0])
        else:
            df = pd.DataFrame([values])

        # Parse checks
        check_list = [c.strip().lower() for c in checks.split(",")]
        if "all" in check_list:
            check_list = ["null", "duplicate", "type"]

        # Run validators
        results = []

        if "null" in check_list:
            validator = NullValidator()
            req_cols = [c.strip() for c in required_columns.split(",")] if required_columns else None
            result = validator.validate(df, required_columns=req_cols)
            results.append(result)

        if "duplicate" in check_list:
            validator = DuplicateValidator()
            key_cols = [c.strip() for c in key_columns.split(",")] if key_columns else None
            result = validator.validate(df, key_columns=key_cols)
            results.append(result)

        if "type" in check_list:
            validator = TypeValidator()
            col_types = None
            if column_types:
                col_types = {}
                for pair in column_types.split(","):
                    col, typ = pair.split(":")
                    col_types[col.strip()] = typ.strip()
            result = validator.validate(df, column_types=col_types, strict=strict_types)
            results.append(result)

        # Generate output
        if output_format == OutputFormat.JSON:
            response = {
                "status": "success",
                "data": {
                    "range": range_addr,
                    "total_rows": len(df),
                    "total_columns": len(df.columns),
                    "validations": [
                        {
                            "validator": r.validator_name,
                            "passed": r.passed,
                            "total_issues": r.total_issues,
                            "issue_rate": r.issue_rate,
                            "summary": r.summary,
                            "issues": r.issues,
                            "details": r.details,
                        }
                        for r in results
                    ],
                    "overall_passed": all(r.passed for r in results),
                },
                "command": "data-validate",
                "message": f"Validated {len(df)} rows × {len(df.columns)} columns",
                "version": get_version(),
            }
            print(json.dumps(response, ensure_ascii=False, indent=2))
        else:
            # Rich text output
            console.print(f"\n[bold cyan]Data Validation Report[/bold cyan]")
            console.print(f"Range: {range_addr}")
            console.print(f"Size: {len(df)} rows × {len(df.columns)} columns\n")

            for result in results:
                # Status icon
                status_icon = "✓" if result.passed else "✗"
                status_color = "green" if result.passed else "red"

                console.print(f"\n[bold]{status_icon} {result.validator_name}[/bold]")
                console.print(f"[{status_color}]{result.summary}[/{status_color}]")

                if result.issues:
                    table = Table(show_header=True)

                    # Dynamic columns based on issue type
                    if result.validator_name == "NullValidator":
                        table.add_column("Column")
                        table.add_column("Null Count")
                        table.add_column("Null Rate")
                        table.add_column("Severity")

                        for issue in result.issues[:10]:  # Show first 10
                            table.add_row(
                                issue["column"],
                                str(issue["null_count"]),
                                f"{issue['null_rate']:.1%}",
                                issue["severity"],
                            )

                    elif result.validator_name == "DuplicateValidator":
                        table.add_column("Type")
                        table.add_column("Count")
                        table.add_column("Description")

                        for issue in result.issues[:10]:
                            table.add_row(
                                issue.get("type", ""),
                                str(issue.get("duplicate_count", issue.get("unique_groups", ""))),
                                issue.get("description", "")[:80],
                            )

                    elif result.validator_name == "TypeValidator":
                        table.add_column("Column")
                        table.add_column("Expected")
                        table.add_column("Actual")
                        table.add_column("Errors")

                        for issue in result.issues[:10]:
                            table.add_row(
                                issue.get("column", ""),
                                issue.get("expected_type", issue.get("detected_type", "")),
                                issue.get("actual_type", ""),
                                str(issue.get("error_count", "")),
                            )

                    console.print(table)

            # Overall result
            overall_passed = all(r.passed for r in results)
            if overall_passed:
                console.print("\n[bold green]✓ All validations passed![/bold green]")
            else:
                console.print("\n[bold red]✗ Some validations failed[/bold red]")

    except Exception as e:
        if output_format == OutputFormat.JSON:
            error_response = {
                "status": "error",
                "error": str(e),
                "version": get_version(),
            }
            print(json.dumps(error_response, ensure_ascii=False, indent=2))
        else:
            console.print(f"[red]Error: {e}[/red]")
        raise typer.Exit(1)
