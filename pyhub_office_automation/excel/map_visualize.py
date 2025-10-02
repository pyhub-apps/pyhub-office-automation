"""
Map visualization CLI command (Issue #72 Phase 3)

Create interactive HTML maps from Seoul district data using Python (no Excel required).
"""

import json
from enum import Enum
from pathlib import Path

import pandas as pd
import typer
from rich.console import Console

from pyhub_office_automation.version import get_version

from .map_visualizer import MapVisualizer

console = Console()


class OutputFormat(str, Enum):
    """Output format options"""

    JSON = "json"
    TEXT = "text"


class MapType(str, Enum):
    """Map visualization type"""

    CHOROPLETH = "choropleth"  # Color-coded regions
    MARKER = "marker"  # Pin markers


class ColorScheme(str, Enum):
    """Color schemes for choropleth maps"""

    YLORD = "YlOrRd"  # Yellow-Orange-Red
    YLGNBU = "YlGnBu"  # Yellow-Green-Blue
    RDYLGN = "RdYlGn"  # Red-Yellow-Green


def map_visualize(
    data_file: str = typer.Option(..., "--data-file", help="CSV/JSON file with district data"),
    value_column: str = typer.Option("value", "--value-column", help="Column name for values"),
    location_column: str = typer.Option("location", "--location-column", help="Column name for locations"),
    output_file: str = typer.Option("seoul_map.html", "--output-file", help="Output HTML file path"),
    map_type: MapType = typer.Option(MapType.CHOROPLETH, "--map-type", help="Map visualization type"),
    title: str = typer.Option("Seoul District Map", "--title", help="Map title"),
    color_scheme: ColorScheme = typer.Option(ColorScheme.YLORD, "--color-scheme", help="Color scheme (choropleth only)"),
    validate_only: bool = typer.Option(False, "--validate-only", help="Only validate data without creating map"),
    output_format: OutputFormat = typer.Option(OutputFormat.TEXT, "--format", help="Output format (json/text)"),
):
    """
    Create interactive map visualization from Seoul district data

    Generates HTML map using Python folium library - no Excel required.
    Works with CSV or JSON input files containing location and value data.

    \\b
    Input Data Format:
      CSV: location,value
           강남구,100
           서초구,85
           ...

      JSON: {"강남구": 100, "서초구": 85, ...}
      or: [{"location": "강남구", "value": 100}, ...]

    \\b
    Examples:
      # Create choropleth map from CSV
      oa excel map-visualize --data-file sales.csv --value-column sales

      # Create marker map with custom title
      oa excel map-visualize --data-file data.json --map-type marker --title "Population"

      # Validate data only
      oa excel map-visualize --data-file sales.csv --validate-only

      # JSON output for AI agents
      oa excel map-visualize --data-file data.csv --format json
    """
    try:
        visualizer = MapVisualizer()

        # Load data
        data_path = Path(data_file)
        if not data_path.exists():
            error_msg = f"Data file not found: {data_file}"
            if output_format == OutputFormat.JSON:
                response = {"status": "error", "error": error_msg, "version": get_version()}
                print(json.dumps(response, ensure_ascii=False, indent=2))
            else:
                console.print(f"[red]Error: {error_msg}[/red]")
            raise typer.Exit(1)

        # Read data based on file extension
        if data_path.suffix.lower() == ".csv":
            df = pd.read_csv(data_path)
            data = df
        elif data_path.suffix.lower() == ".json":
            with open(data_path, "r", encoding="utf-8") as f:
                json_data = json.load(f)

            # Handle different JSON formats
            if isinstance(json_data, dict):
                # Direct dict format
                data = json_data
            elif isinstance(json_data, list):
                # List of dicts format
                df = pd.DataFrame(json_data)
                data = df
            else:
                raise ValueError("Unsupported JSON format")
        else:
            raise ValueError(f"Unsupported file format: {data_path.suffix}")

        # Validate data
        validation_result = visualizer.validate_data(data)

        if validate_only:
            if output_format == OutputFormat.JSON:
                response = {
                    "status": "success",
                    "data": validation_result,
                    "command": "map-visualize",
                    "message": "Data validation completed",
                    "version": get_version(),
                }
                print(json.dumps(response, ensure_ascii=False, indent=2))
            else:
                console.print("\n[bold cyan]Data Validation Results[/bold cyan]\n")
                console.print(f"Total Locations: {validation_result['total_locations']}")
                console.print(f"[green]Matched: {validation_result['matched_count']}[/green]")
                console.print(f"[red]Unmatched: {validation_result['unmatched_count']}[/red]")

                if validation_result["matched"]:
                    console.print("\n[green]Matched Locations:[/green]")
                    for item in validation_result["matched"][:10]:
                        console.print(f"  ✓ {item['input']} → {item['matched']}")

                if validation_result["unmatched"]:
                    console.print("\n[red]Unmatched Locations:[/red]")
                    for item in validation_result["unmatched"][:10]:
                        console.print(f"  ✗ {item['input']}")
                        if item["suggestions"]:
                            console.print(f"    Suggestions: {', '.join(item['suggestions'])}")

            return

        # Warn about unmatched data
        if validation_result["unmatched_count"] > 0:
            if output_format == OutputFormat.TEXT:
                console.print(
                    f"\n[yellow]Warning: {validation_result['unmatched_count']} locations could not be matched[/yellow]"
                )

        # Create map
        if map_type == MapType.CHOROPLETH:
            output_path = visualizer.create_choropleth_map(
                data=data,
                value_column=value_column if isinstance(data, pd.DataFrame) else None,
                location_column=location_column,
                output_file=output_file,
                title=title,
                color_scheme=color_scheme.value,
            )
        else:  # MARKER
            output_path = visualizer.create_marker_map(
                data=data,
                value_column=value_column if isinstance(data, pd.DataFrame) else None,
                location_column=location_column,
                output_file=output_file,
                title=title,
            )

        # Output result
        if output_format == OutputFormat.JSON:
            response = {
                "status": "success",
                "data": {
                    "output_file": output_path,
                    "map_type": map_type.value,
                    "validation": validation_result,
                },
                "command": "map-visualize",
                "message": f"Map created successfully: {output_path}",
                "version": get_version(),
            }
            print(json.dumps(response, ensure_ascii=False, indent=2))
        else:
            console.print(f"\n[bold green]✓ Map created successfully![/bold green]")
            console.print(f"Output file: {output_path}")
            console.print(f"Map type: {map_type.value}")
            console.print(f"Matched locations: {validation_result['matched_count']}")
            console.print(f"\nOpen the file in your browser to view the interactive map.")

    except Exception as e:
        if output_format == OutputFormat.JSON:
            error_response = {"status": "error", "error": str(e), "version": get_version()}
            print(json.dumps(error_response, ensure_ascii=False, indent=2))
        else:
            console.print(f"[red]Error: {e}[/red]")
        raise typer.Exit(1)
