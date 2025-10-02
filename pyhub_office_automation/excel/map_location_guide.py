"""
Map Chart location name guidance command (Issue #72 Phase 2)

Provides user-friendly guidance for Excel Map Chart location naming conventions.
"""

import json
from enum import Enum

import typer
from rich.console import Console
from rich.table import Table

from pyhub_office_automation.version import get_version

from .location_converter import LocationConverter

console = Console()


class OutputFormat(str, Enum):
    """Output format options"""

    JSON = "json"
    TEXT = "text"


def map_location_guide(
    region: str = typer.Option("seoul", "--region", help="Region name (seoul, busan, etc.)"),
    show_all: bool = typer.Option(False, "--show-all", help="Show all district names"),
    test_names: str = typer.Option(None, "--test", help="Test location names (comma-separated)"),
    output_format: OutputFormat = typer.Option(OutputFormat.TEXT, "--format", help="Output format (json/text)"),
):
    """
        Get location name format guidance for Excel Map Chart

        Displays proper naming conventions, accepted formats, and validation tips
        for creating Map Charts with Korean location data.

        \b
        Supported Regions:
          " seoul: Seoul 25 districts (�l, Gangnam, etc.)
          " (more regions coming in future updates)

        \b
        Examples:
          # Show Seoul district guidance
          oa excel map-location-guide --region seoul

          # Show all 25 district names
          oa excel map-location-guide --region seoul --show-all

          # Test specific location names
          oa excel map-location-guide --test "�l,Gangnam-gu,
    � �"

          # JSON output for AI agents
          oa excel map-location-guide --region seoul --format json
    """
    try:
        converter = LocationConverter()
        guidance = converter.get_guidance(region)

        # Check if region is supported
        if "error" in guidance:
            if output_format == OutputFormat.JSON:
                error_response = {
                    "status": "error",
                    "error": guidance["error"],
                    "supported_regions": guidance["supported_regions"],
                    "version": get_version(),
                }
                print(json.dumps(error_response, ensure_ascii=False, indent=2))
            else:
                console.print(f"[red]Error: {guidance['error']}[/red]")
                console.print(f"Supported regions: {', '.join(guidance['supported_regions'])}")
            raise typer.Exit(1)

        # Test mode
        if test_names:
            test_results = []
            for name in test_names.split(","):
                name = name.strip()
                result = converter.convert_seoul_district(name)
                test_results.append(
                    {
                        "input": name,
                        "matched": result.matched,
                        "status": result.status,
                        "confidence": result.confidence,
                        "suggestions": result.suggestions[:3],
                    }
                )

            if output_format == OutputFormat.JSON:
                response = {
                    "status": "success",
                    "data": {"test_results": test_results},
                    "command": "map-location-guide",
                    "message": f"Tested {len(test_results)} location names",
                    "version": get_version(),
                }
                print(json.dumps(response, ensure_ascii=False, indent=2))
            else:
                console.print("\n[bold cyan]Location Name Test Results[/bold cyan]\n")
                for tr in test_results:
                    status_color = (
                        "green" if tr["status"] in ["exact", "converted"] else "yellow" if tr["status"] == "fuzzy" else "red"
                    )
                    console.print(f"Input: [bold]{tr['input']}[/bold]")
                    console.print(f"  Status: [{status_color}]{tr['status']}[/{status_color}]")
                    console.print(f"  Matched: {tr['matched'] or 'None'}")
                    console.print(f"  Confidence: {tr['confidence']:.0%}")
                    if tr["suggestions"]:
                        console.print(f"  Suggestions: {', '.join(tr['suggestions'])}")
                    console.print()

            return

        # Normal guidance mode
        if output_format == OutputFormat.JSON:
            response = {
                "status": "success",
                "data": guidance,
                "command": "map-location-guide",
                "message": f"Location guidance for {region}",
                "version": get_version(),
            }
            print(json.dumps(response, ensure_ascii=False, indent=2))
        else:
            # Rich text output
            console.print(f"\n[bold cyan]Excel Map Chart - {guidance['region']} Location Name Guide[/bold cyan]\n")

            # Basic info
            console.print(f"[bold]Total Districts:[/bold] {guidance['total_districts']}")
            console.print()

            # Correct formats
            console.print("[bold green]Correct Formats (Excel recognizes these):[/bold green]")
            for fmt in guidance["correct_formats"]:
                console.print(f"  • {fmt}")
            console.print()

            # Accepted inputs
            console.print("[bold yellow]Accepted Inputs (auto-converted):[/bold yellow]")
            for inp in guidance["accepted_inputs"]:
                console.print(f"  • {inp}")
            console.print()

            # Tips
            console.print("[bold cyan]Tips & Requirements:[/bold cyan]")
            for i, tip in enumerate(guidance["tips"], 1):
                console.print(f"  {i}. {tip}")
            console.print()

            # Show all districts if requested
            if show_all:
                console.print("[bold magenta]All 25 Seoul Districts (Excel Format):[/bold magenta]")
                districts = guidance["all_districts"]
                # Display in 3 columns
                for i in range(0, len(districts), 3):
                    row = districts[i : i + 3]
                    console.print("  " + "    ".join(row))
                console.print()

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
