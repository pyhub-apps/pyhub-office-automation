"""
Batch script executor for Office Automation

Executes sequences of shell commands from script files (.oas format).

Phase 1: Basic execution without variables or control flow
"""

import shlex
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import List, Optional

import typer
from rich.console import Console
from rich.progress import Progress, SpinnerColumn, TextColumn
from typer.testing import CliRunner

console = Console()
runner = CliRunner()


@dataclass
class BatchLine:
    """Single line in batch script"""

    line_number: int
    content: str
    command: Optional[str] = None
    args: List[str] = field(default_factory=list)
    is_comment: bool = False
    is_directive: bool = False
    is_empty: bool = False


@dataclass
class LineResult:
    """Single line execution result"""

    line_number: int
    command: str
    success: bool
    output: str = ""
    error: Optional[str] = None
    duration_ms: int = 0


@dataclass
class BatchResult:
    """Overall batch execution result"""

    success: bool
    executed_lines: int
    skipped_lines: int
    failed_lines: int
    total_duration_ms: int
    log: List[LineResult]
    start_time: datetime
    end_time: datetime


def parse_script(script_path: str) -> List[BatchLine]:
    """
    Parse batch script file into BatchLine objects

    Supports:
    - Comments (# ...)
    - Empty lines
    - Shell commands
    """
    lines = []
    path = Path(script_path)

    if not path.exists():
        raise FileNotFoundError(f"Script file not found: {script_path}")

    with open(path, "r", encoding="utf-8") as f:
        for line_num, line in enumerate(f, 1):
            line = line.rstrip()

            # Empty line
            if not line:
                lines.append(BatchLine(line_number=line_num, content=line, is_empty=True))
                continue

            # Comment
            if line.strip().startswith("#"):
                lines.append(BatchLine(line_number=line_num, content=line, is_comment=True))
                continue

            # Directive (Phase 2 - not implemented yet)
            if line.strip().startswith("@"):
                lines.append(
                    BatchLine(
                        line_number=line_num,
                        content=line,
                        is_directive=True,
                    )
                )
                console.print(f"[yellow]Warning: Directives not supported yet (line {line_num}): {line}[/yellow]")
                continue

            # Shell command
            try:
                tokens = shlex.split(line)
                if tokens:
                    lines.append(
                        BatchLine(
                            line_number=line_num,
                            content=line,
                            command=tokens[0],
                            args=tokens[1:],
                        )
                    )
            except ValueError as e:
                console.print(f"[red]Parse error at line {line_num}: {e}[/red]")
                lines.append(
                    BatchLine(
                        line_number=line_num,
                        content=line,
                        is_comment=True,  # Treat as comment to skip
                    )
                )

    return lines


def execute_shell_command(line: BatchLine, mode: str = "unified") -> LineResult:
    """
    Execute a single shell command

    Args:
        line: Parsed batch line
        mode: "unified", "excel", or "ppt"

    Returns:
        Execution result
    """
    from pyhub_office_automation.cli.main import app as main_app

    start_time = datetime.now()

    try:
        # Build command arguments
        if mode == "unified":
            # Unified shell commands
            cmd_args = [line.command] + line.args
        else:
            # Mode-specific commands (future: route to excel/ppt)
            cmd_args = [line.command] + line.args

        # Execute command
        result = runner.invoke(main_app, cmd_args)

        duration_ms = int((datetime.now() - start_time).total_seconds() * 1000)

        if result.exit_code == 0:
            return LineResult(
                line_number=line.line_number,
                command=line.content,
                success=True,
                output=result.stdout,
                duration_ms=duration_ms,
            )
        else:
            return LineResult(
                line_number=line.line_number,
                command=line.content,
                success=False,
                output=result.stdout,
                error=f"Exit code: {result.exit_code}",
                duration_ms=duration_ms,
            )

    except Exception as e:
        duration_ms = int((datetime.now() - start_time).total_seconds() * 1000)
        return LineResult(
            line_number=line.line_number,
            command=line.content,
            success=False,
            error=str(e),
            duration_ms=duration_ms,
        )


def batch_run(
    script_path: str,
    dry_run: bool = False,
    verbose: bool = False,
    continue_on_error: bool = False,
    log_file: Optional[str] = None,
):
    """
    Execute batch script

    Args:
        script_path: Path to .oas script file
        dry_run: If True, parse but don't execute
        verbose: Show detailed output
        continue_on_error: Continue execution even if commands fail
        log_file: Path to log file
    """
    console.print(f"\n[bold cyan]Batch Script Execution[/bold cyan]: {script_path}\n")

    start_time = datetime.now()

    # Parse script
    try:
        lines = parse_script(script_path)
    except Exception as e:
        console.print(f"[red]Failed to parse script: {e}[/red]")
        return

    total_lines = len(lines)
    executable_lines = [l for l in lines if not (l.is_comment or l.is_empty or l.is_directive)]

    console.print(f"Total lines: {total_lines}")
    console.print(f"Executable lines: {len(executable_lines)}")
    console.print(f"Comments/Empty: {total_lines - len(executable_lines)}\n")

    if dry_run:
        console.print("[yellow]Dry-run mode - commands will not be executed[/yellow]\n")
        for line in executable_lines:
            console.print(f"[dim]Line {line.line_number}:[/dim] {line.content}")
        return

    # Execute script
    results = []
    failed_count = 0

    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        console=console,
    ) as progress:
        task = progress.add_task("[cyan]Executing commands...", total=len(executable_lines))

        for line in lines:
            # Skip non-executable lines
            if line.is_comment or line.is_empty or line.is_directive:
                continue

            if verbose:
                console.print(f"\n[dim]Line {line.line_number}:[/dim] {line.content}")

            # Execute command
            result = execute_shell_command(line)
            results.append(result)

            if verbose:
                if result.success:
                    console.print("[green]✓ Success[/green]")
                    if result.output:
                        console.print(result.output)
                else:
                    console.print(f"[red]✗ Failed: {result.error}[/red]")
                    if result.output:
                        console.print(result.output)

            if not result.success:
                failed_count += 1
                if not continue_on_error:
                    console.print(f"\n[red]Execution stopped at line {line.line_number} due to error[/red]")
                    break

            progress.update(task, advance=1)

    end_time = datetime.now()
    total_duration_ms = int((end_time - start_time).total_seconds() * 1000)

    # Summary
    batch_result = BatchResult(
        success=failed_count == 0,
        executed_lines=len(results),
        skipped_lines=total_lines - len(executable_lines),
        failed_lines=failed_count,
        total_duration_ms=total_duration_ms,
        log=results,
        start_time=start_time,
        end_time=end_time,
    )

    console.print("\n" + "=" * 60)
    console.print("[bold cyan]Execution Summary[/bold cyan]")
    console.print("=" * 60)
    console.print(f"Total lines executed: {batch_result.executed_lines}")
    console.print(f"Successful: {batch_result.executed_lines - failed_count}")
    console.print(f"Failed: {failed_count}")
    console.print(f"Skipped (comments/empty): {batch_result.skipped_lines}")
    console.print(f"Total duration: {total_duration_ms}ms")

    if batch_result.success:
        console.print("\n[bold green]✓ Batch execution completed successfully![/bold green]")
    else:
        console.print(f"\n[bold red]✗ Batch execution failed with {failed_count} error(s)[/bold red]")

    # Write log file
    if log_file:
        write_log_file(log_file, batch_result, script_path)
        console.print(f"\nLog written to: {log_file}")


def write_log_file(log_path: str, result: BatchResult, script_path: str):
    """Write execution log to file"""
    with open(log_path, "w", encoding="utf-8") as f:
        f.write("=" * 60 + "\n")
        f.write("Batch Execution Log\n")
        f.write("=" * 60 + "\n")
        f.write(f"Script: {script_path}\n")
        f.write(f"Start Time: {result.start_time.isoformat()}\n")
        f.write(f"End Time: {result.end_time.isoformat()}\n")
        f.write(f"Duration: {result.total_duration_ms}ms\n")
        f.write(f"Success: {result.success}\n")
        f.write("\n")

        f.write("Execution Details:\n")
        f.write("-" * 60 + "\n")

        for line_result in result.log:
            status = "✓ SUCCESS" if line_result.success else "✗ FAILED"
            f.write(f"\nLine {line_result.line_number}: {status}\n")
            f.write(f"Command: {line_result.command}\n")
            f.write(f"Duration: {line_result.duration_ms}ms\n")

            if line_result.output:
                f.write(f"Output:\n{line_result.output}\n")

            if line_result.error:
                f.write(f"Error: {line_result.error}\n")

            f.write("-" * 60 + "\n")

        f.write("\nSummary:\n")
        f.write(f"Total Executed: {result.executed_lines}\n")
        f.write(f"Failed: {result.failed_lines}\n")
        f.write(f"Skipped: {result.skipped_lines}\n")
