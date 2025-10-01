"""
Batch script executor for Office Automation

Executes sequences of shell commands from script files (.oas format).

Phase 3: Control flow support (@if, @foreach, @while)
"""

import shlex
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

import typer
from rich.console import Console
from rich.progress import Progress, SpinnerColumn, TextColumn
from typer.testing import CliRunner

from .control_flow import (
    ConditionEvaluator,
    parse_elif_condition,
    parse_foreach_loop,
    parse_if_condition,
    parse_list_expression,
    parse_onerror_directive,
    parse_while_condition,
)
from .variables import (
    VariableManager,
    parse_echo_directive,
    parse_export_directive,
    parse_set_directive,
    parse_unset_directive,
)

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

            # Directive (Phase 2 - variable directives supported)
            if line.strip().startswith("@"):
                lines.append(
                    BatchLine(
                        line_number=line_num,
                        content=line,
                        is_directive=True,
                    )
                )
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


def find_matching_endif(lines: List[BatchLine], start_idx: int) -> int:
    """
    Find matching @endif for @if directive

    Args:
        lines: List of batch lines
        start_idx: Index of @if line

    Returns:
        Index of matching @endif

    Raises:
        ValueError: If no matching @endif found
    """
    depth = 1
    for i in range(start_idx + 1, len(lines)):
        line = lines[i]
        if line.is_directive:
            content = line.content.strip()
            if content.startswith("@if "):
                depth += 1
            elif content == "@endif":
                depth -= 1
                if depth == 0:
                    return i

    raise ValueError(f"No matching @endif found for @if at line {lines[start_idx].line_number}")


def find_matching_endforeach(lines: List[BatchLine], start_idx: int) -> int:
    """
    Find matching @endforeach for @foreach directive

    Args:
        lines: List of batch lines
        start_idx: Index of @foreach line

    Returns:
        Index of matching @endforeach

    Raises:
        ValueError: If no matching @endforeach found
    """
    depth = 1
    for i in range(start_idx + 1, len(lines)):
        line = lines[i]
        if line.is_directive:
            content = line.content.strip()
            if content.startswith("@foreach "):
                depth += 1
            elif content == "@endforeach":
                depth -= 1
                if depth == 0:
                    return i

    raise ValueError(f"No matching @endforeach found for @foreach at line {lines[start_idx].line_number}")


def find_elif_else_blocks(lines: List[BatchLine], if_idx: int, endif_idx: int) -> tuple[List[tuple[int, str]], Optional[int]]:
    """
    Find all @elif and @else blocks within @if...@endif

    Args:
        lines: List of batch lines
        if_idx: Index of @if line
        endif_idx: Index of @endif line

    Returns:
        Tuple of (elif_blocks, else_idx)
        elif_blocks: List of (line_idx, condition)
        else_idx: Index of @else line or None
    """
    elif_blocks = []
    else_idx = None
    depth = 0

    for i in range(if_idx + 1, endif_idx):
        line = lines[i]
        if line.is_directive:
            content = line.content.strip()

            if content.startswith("@if "):
                depth += 1
            elif content == "@endif":
                depth -= 1
            elif depth == 0:  # Same level as our @if
                if content.startswith("@elif "):
                    condition = parse_elif_condition(content)
                    elif_blocks.append((i, condition))
                elif content == "@else":
                    else_idx = i

    return elif_blocks, else_idx


def find_matching_endtry(lines: List[BatchLine], start_idx: int) -> int:
    """
    Find matching @endtry for @try directive

    Args:
        lines: List of batch lines
        start_idx: Index of @try line

    Returns:
        Index of matching @endtry

    Raises:
        ValueError: If no matching @endtry found
    """
    depth = 1
    for i in range(start_idx + 1, len(lines)):
        line = lines[i]
        if line.is_directive:
            content = line.content.strip()
            if content == "@try":
                depth += 1
            elif content == "@endtry":
                depth -= 1
                if depth == 0:
                    return i

    raise ValueError(f"No matching @endtry found for @try at line {lines[start_idx].line_number}")


def find_catch_finally_blocks(lines: List[BatchLine], try_idx: int, endtry_idx: int) -> tuple[Optional[int], Optional[int]]:
    """
    Find @catch and @finally blocks within @try...@endtry

    Args:
        lines: List of batch lines
        try_idx: Index of @try line
        endtry_idx: Index of @endtry line

    Returns:
        Tuple of (catch_idx, finally_idx)
        catch_idx: Index of @catch line or None
        finally_idx: Index of @finally line or None
    """
    catch_idx = None
    finally_idx = None
    depth = 0

    for i in range(try_idx + 1, endtry_idx):
        line = lines[i]
        if line.is_directive:
            content = line.content.strip()

            if content == "@try":
                depth += 1
            elif content == "@endtry":
                depth -= 1
            elif depth == 0:  # Same level as our @try
                if content == "@catch":
                    catch_idx = i
                elif content == "@finally":
                    finally_idx = i

    return catch_idx, finally_idx


def execute_directive(line: BatchLine, var_manager: VariableManager) -> LineResult:
    """
    Execute a directive line (@set, @unset, @echo, @export)

    Args:
        line: Parsed batch line with directive
        var_manager: Variable manager instance

    Returns:
        Execution result
    """
    start_time = datetime.now()
    content = line.content.strip()

    try:
        # @set VAR = value
        if content.startswith("@set "):
            name, value = parse_set_directive(content)
            # Resolve variables in value before setting
            resolved_value = var_manager.resolve(value)
            var_manager.set(name, resolved_value)
            output = f"Variable set: {name} = {resolved_value}"

        # @unset VAR
        elif content.startswith("@unset "):
            name = parse_unset_directive(content)
            var_manager.unset(name)
            output = f"Variable unset: {name}"

        # @echo message
        elif content.startswith("@echo "):
            message = parse_echo_directive(content)
            resolved_message = var_manager.resolve(message)
            console.print(f"[cyan]{resolved_message}[/cyan]")
            output = resolved_message

        # @export VAR = value
        elif content.startswith("@export "):
            name, value = parse_export_directive(content)
            resolved_value = var_manager.resolve(value)
            var_manager.export(name, resolved_value)
            output = f"Variable exported: {name} = {resolved_value}"

        else:
            # Unknown directive - warn but don't fail
            output = f"Unknown directive (skipped): {content}"
            console.print(f"[yellow]Warning: {output}[/yellow]")

        duration_ms = int((datetime.now() - start_time).total_seconds() * 1000)

        return LineResult(
            line_number=line.line_number,
            command=line.content,
            success=True,
            output=output,
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


def execute_shell_command(line: BatchLine, var_manager: VariableManager, mode: str = "unified") -> LineResult:
    """
    Execute a single shell command with variable substitution

    Args:
        line: Parsed batch line
        var_manager: Variable manager for variable resolution
        mode: "unified", "excel", or "ppt"

    Returns:
        Execution result
    """
    from pyhub_office_automation.cli.main import app as main_app

    start_time = datetime.now()

    try:
        # Resolve variables in command and arguments
        resolved_command = var_manager.resolve(line.command)
        resolved_args = [var_manager.resolve(arg) for arg in line.args]

        # Build command arguments
        if mode == "unified":
            # Unified shell commands
            cmd_args = [resolved_command] + resolved_args
        else:
            # Mode-specific commands (future: route to excel/ppt)
            cmd_args = [resolved_command] + resolved_args

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


def execute_lines(
    lines: List[BatchLine],
    var_manager: VariableManager,
    start_idx: int = 0,
    end_idx: Optional[int] = None,
    verbose: bool = False,
    continue_on_error: bool = False,
) -> tuple[List[LineResult], int]:
    """
    Execute a range of lines with control flow support

    Args:
        lines: List of all batch lines
        var_manager: Variable manager
        start_idx: Start index (inclusive)
        end_idx: End index (exclusive), None means till end
        verbose: Show detailed output
        continue_on_error: Continue on errors

    Returns:
        Tuple of (results, next_index_to_process)
    """
    if end_idx is None:
        end_idx = len(lines)

    results = []
    i = start_idx
    evaluator = ConditionEvaluator(var_manager)

    while i < end_idx:
        line = lines[i]

        # Skip comments and empty lines
        if line.is_comment or line.is_empty:
            i += 1
            continue

        # Handle control flow directives
        if line.is_directive:
            content = line.content.strip()

            # @if statement
            if content.startswith("@if "):
                endif_idx = find_matching_endif(lines, i)
                elif_blocks, else_idx = find_elif_else_blocks(lines, i, endif_idx)

                # Evaluate @if condition
                condition = parse_if_condition(content)
                if evaluator.evaluate(condition):
                    # Execute if block (until first elif/else or endif)
                    block_end = elif_blocks[0][0] if elif_blocks else (else_idx if else_idx else endif_idx)
                    block_results, _ = execute_lines(lines, var_manager, i + 1, block_end, verbose, continue_on_error)
                    results.extend(block_results)
                else:
                    # Try elif blocks
                    executed = False
                    for elif_idx, elif_cond in elif_blocks:
                        if evaluator.evaluate(elif_cond):
                            # Find next elif or else or endif
                            next_elif_idx = None
                            for next_idx, _ in elif_blocks:
                                if next_idx > elif_idx:
                                    next_elif_idx = next_idx
                                    break
                            block_end = next_elif_idx if next_elif_idx else (else_idx if else_idx else endif_idx)
                            block_results, _ = execute_lines(
                                lines, var_manager, elif_idx + 1, block_end, verbose, continue_on_error
                            )
                            results.extend(block_results)
                            executed = True
                            break

                    # Execute else block if no condition was true
                    if not executed and else_idx is not None:
                        block_results, _ = execute_lines(
                            lines, var_manager, else_idx + 1, endif_idx, verbose, continue_on_error
                        )
                        results.extend(block_results)

                # Skip to after @endif
                i = endif_idx + 1
                continue

            # @foreach loop
            elif content.startswith("@foreach "):
                endforeach_idx = find_matching_endforeach(lines, i)
                var_name, list_expr = parse_foreach_loop(content)
                items = parse_list_expression(list_expr, var_manager)

                # Execute loop body for each item
                for loop_idx, item in enumerate(items):
                    # Set loop variable
                    var_manager.set(var_name, item)
                    var_manager.set("__LOOP_INDEX__", str(loop_idx))

                    # Execute loop body
                    loop_results, _ = execute_lines(lines, var_manager, i + 1, endforeach_idx, verbose, continue_on_error)
                    results.extend(loop_results)

                # Clean up loop variables
                var_manager.unset(var_name)
                var_manager.unset("__LOOP_INDEX__")

                # Skip to after @endforeach
                i = endforeach_idx + 1
                continue

            # @try/@catch/@finally block
            elif content == "@try":
                endtry_idx = find_matching_endtry(lines, i)
                catch_idx, finally_idx = find_catch_finally_blocks(lines, i, endtry_idx)

                error_occurred = False
                try_results = []

                # Execute try block
                try_end = catch_idx if catch_idx else (finally_idx if finally_idx else endtry_idx)
                try:
                    try_results, _ = execute_lines(lines, var_manager, i + 1, try_end, verbose, True)  # Force continue in try
                    # Check if any command failed
                    error_occurred = any(not r.success for r in try_results)
                    results.extend(try_results)
                except Exception as e:
                    error_occurred = True
                    console.print(f"[red]Exception in try block: {e}[/red]")

                # Execute catch block if error occurred
                if error_occurred and catch_idx is not None:
                    catch_end = finally_idx if finally_idx else endtry_idx
                    catch_results, _ = execute_lines(lines, var_manager, catch_idx + 1, catch_end, verbose, continue_on_error)
                    results.extend(catch_results)

                # Always execute finally block
                if finally_idx is not None:
                    finally_results, _ = execute_lines(
                        lines, var_manager, finally_idx + 1, endtry_idx, verbose, continue_on_error
                    )
                    results.extend(finally_results)

                # Skip to after @endtry
                i = endtry_idx + 1
                continue

            # Skip @endif, @endforeach, @elif, @else, @catch, @finally, @endtry (handled by parent)
            elif content in ("@endif", "@endforeach", "@else", "@catch", "@finally", "@endtry") or content.startswith(
                "@elif "
            ):
                i += 1
                continue

            # Other directives (variable management)
            else:
                result = execute_directive(line, var_manager)
                results.append(result)

                if not result.success and not continue_on_error:
                    return results, i + 1

        # Regular shell command
        else:
            result = execute_shell_command(line, var_manager)
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

            if not result.success and not continue_on_error:
                return results, i + 1

        i += 1

    return results, i


def batch_run(
    script_path: str,
    dry_run: bool = False,
    verbose: bool = False,
    continue_on_error: bool = False,
    log_file: Optional[str] = None,
    variables: Optional[Dict[str, str]] = None,
):
    """
    Execute batch script with variable and control flow support

    Args:
        script_path: Path to .oas script file
        dry_run: If True, parse but don't execute
        verbose: Show detailed output
        continue_on_error: Continue execution even if commands fail
        log_file: Path to log file
        variables: Initial variables (from --set CLI options)
    """
    console.print(f"\n[bold cyan]Batch Script Execution[/bold cyan]: {script_path}\n")

    start_time = datetime.now()

    # Initialize variable manager with CLI variables
    var_manager = VariableManager(initial_vars=variables or {})

    # Parse script
    try:
        lines = parse_script(script_path)
    except Exception as e:
        console.print(f"[red]Failed to parse script: {e}[/red]")
        return

    total_lines = len(lines)
    directive_lines = [l for l in lines if l.is_directive]
    executable_lines = [l for l in lines if not (l.is_comment or l.is_empty)]

    console.print(f"Total lines: {total_lines}")
    console.print(f"Executable lines: {len(executable_lines)}")
    console.print(f"  - Directives: {len(directive_lines)}")
    console.print(f"  - Commands: {len(executable_lines) - len(directive_lines)}")
    console.print(f"Comments/Empty: {total_lines - len(executable_lines)}\n")

    if variables:
        console.print("[bold]Initial Variables:[/bold]")
        for name, value in variables.items():
            console.print(f"  {name} = {value}")
        console.print()

    if dry_run:
        console.print("[yellow]Dry-run mode - commands will not be executed[/yellow]\n")
        for line in executable_lines:
            if line.is_directive:
                console.print(f"[dim]Line {line.line_number} [Directive]:[/dim] {line.content}")
            else:
                # Resolve variables for preview
                resolved_cmd = var_manager.resolve(line.command)
                resolved_args = [var_manager.resolve(arg) for arg in line.args]
                resolved_line = f"{resolved_cmd} {' '.join(resolved_args)}"
                console.print(f"[dim]Line {line.line_number}:[/dim] {resolved_line}")
        return

    # Execute script with control flow support
    results = []
    failed_count = 0

    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        console=console,
    ) as progress:
        task = progress.add_task("[cyan]Executing commands...", total=len(executable_lines))

        # Execute all lines with control flow
        results, _ = execute_lines(lines, var_manager, 0, None, verbose, continue_on_error)
        failed_count = sum(1 for r in results if not r.success)

        progress.update(task, advance=len(executable_lines))

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
