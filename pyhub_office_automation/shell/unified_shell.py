"""
Unified Shell Mode - Excel and PowerPoint integration

Provides a unified REPL interface for managing both Excel and PowerPoint
within a single shell session, allowing seamless context switching and
integrated workflows.

Issue #87: Unified Shell (Excel + PowerPoint)
"""

import shlex
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional

import typer
from prompt_toolkit import PromptSession
from prompt_toolkit.completion import Completer, Completion
from prompt_toolkit.history import FileHistory
from rich.console import Console
from typer.testing import CliRunner

console = Console()
runner = CliRunner()


@dataclass
class UnifiedShellContext:
    """Unified shell session context for Excel and PowerPoint"""

    mode: str = "none"  # "excel", "ppt", "none"

    # Excel context
    excel_workbook_path: Optional[str] = None
    excel_workbook_name: Optional[str] = None
    excel_sheet: Optional[str] = None
    excel_app: Optional[object] = None

    # PowerPoint context
    ppt_presentation_path: Optional[str] = None
    ppt_presentation_name: Optional[str] = None
    ppt_slide_number: Optional[int] = None
    ppt_prs: Optional[object] = None

    def get_prompt_text(self) -> str:
        """Generate prompt text based on current mode"""
        if self.mode == "excel":
            wb_name = self.excel_workbook_name or "None"
            sheet = self.excel_sheet or "None"
            return f"[OA Shell: Excel {wb_name} > Sheet {sheet}] > "
        elif self.mode == "ppt":
            prs_name = self.ppt_presentation_name or "None"
            slide = str(self.ppt_slide_number) if self.ppt_slide_number else "None"
            return f"[OA Shell: PPT {prs_name} > Slide {slide}] > "
        else:
            return "[OA Shell] > "


# Excel commands mapping (from excel_shell.py)
EXCEL_COMMANDS = [
    "sheet-activate",
    "sheet-add",
    "sheet-delete",
    "sheet-rename",
    "workbook-create",
    "workbook-open",
    "workbook-list",
    "workbook-info",
    "range-read",
    "range-write",
    "table-read",
    "table-write",
    "table-list",
    "table-analyze",
    "metadata-generate",
    "chart-add",
    "chart-pivot-create",
    "chart-list",
    "chart-configure",
    "chart-position",
    "chart-export",
    "chart-delete",
    "sheets",  # Shell-specific alias
]

# PowerPoint commands mapping (from ppt_shell.py)
PPT_COMMANDS = [
    "presentation-create",
    "presentation-open",
    "presentation-save",
    "presentation-list",
    "presentation-info",
    "slide-list",
    "slide-add",
    "slide-delete",
    "slide-duplicate",
    "slide-copy",
    "slide-reorder",
    "content-add-text",
    "content-add-image",
    "content-add-shape",
    "content-add-table",
    "content-add-chart",
    "content-add-video",
    "content-add-smartart",
    "content-add-excel-chart",
    "content-add-audio",
    "content-add-equation",
    "content-update",
    "layout-list",
    "layout-apply",
    "template-apply",
    "theme-apply",
    "export-pdf",
    "export-images",
    "export-notes",
    "slideshow-start",
    "slideshow-control",
    "run-macro",
    "animation-add",
    "slides",  # Shell-specific alias
]

# Unified shell commands (available in all modes)
UNIFIED_COMMANDS = ["help", "show", "clear", "exit", "quit", "use", "switch"]


class UnifiedShellCompleter(Completer):
    """Smart autocomplete based on current mode"""

    def __init__(self, context: UnifiedShellContext):
        self.context = context

    def get_completions(self, document, complete_event):
        word = document.get_word_before_cursor()

        # Determine available commands based on mode
        if self.context.mode == "excel":
            available_commands = UNIFIED_COMMANDS + EXCEL_COMMANDS
        elif self.context.mode == "ppt":
            available_commands = UNIFIED_COMMANDS + PPT_COMMANDS
        else:
            # No mode active - only unified commands and mode keywords
            available_commands = UNIFIED_COMMANDS + ["excel", "ppt"]

        # Generate completions
        for cmd in sorted(set(available_commands)):
            if cmd.startswith(word):
                yield Completion(cmd, start_position=-len(word))


def show_unified_help(ctx: UnifiedShellContext):
    """Display comprehensive help based on current mode"""
    console.print("\n[bold cyan]Unified Shell Commands (Always Available):[/bold cyan]")
    console.print("  help              - Show this help message")
    console.print("  show context      - Display current context (Excel + PowerPoint)")
    console.print("  clear             - Clear terminal screen")
    console.print("  exit / quit       - Exit shell")
    console.print("  use excel <file>  - Switch to Excel mode and open file")
    console.print("  use ppt <file>    - Switch to PowerPoint mode and open file")
    console.print("  switch excel      - Switch to Excel mode (if loaded)")
    console.print("  switch ppt        - Switch to PowerPoint mode (if loaded)")

    if ctx.mode == "excel":
        console.print(f"\n[bold green]Excel Commands ({len(EXCEL_COMMANDS)} available):[/bold green]")
        console.print("  sheets, workbook-info, range-read, range-write,")
        console.print("  table-read, table-list, chart-add, sheet-add, sheet-delete, ...")
        console.print("\n  Use Tab for autocomplete or 'oa excel --help' for full list")

    elif ctx.mode == "ppt":
        console.print(f"\n[bold magenta]PowerPoint Commands ({len(PPT_COMMANDS)} available):[/bold magenta]")
        console.print("  slides, presentation-info, content-add-text, content-add-image,")
        console.print("  slide-add, slide-delete, layout-apply, export-pdf, ...")
        console.print("\n  Use Tab for autocomplete or 'oa ppt --help' for full list")

    else:
        console.print("\n[yellow]No active mode. Use 'use excel <file>' or 'use ppt <file>' to start.[/yellow]")

    console.print("\n[dim]Press Tab for command autocomplete[/dim]")
    console.print("[dim]Press ↑/↓ for command history[/dim]\n")


def show_unified_context(ctx: UnifiedShellContext):
    """Display current context for both Excel and PowerPoint"""
    console.print("\n[bold cyan]Current Context:[/bold cyan]")
    console.print(f"  Active Mode: [bold]{ctx.mode.upper()}[/bold]\n")

    # Excel context
    console.print("[bold green]Excel Context:[/bold green]")
    if ctx.excel_workbook_path:
        console.print(f"  Workbook: {ctx.excel_workbook_name}")
        console.print(f"  Path: {ctx.excel_workbook_path}")
        console.print(f"  Active Sheet: {ctx.excel_sheet or 'None'}")
        if ctx.mode != "excel":
            console.print("  [dim](Use 'switch excel' to activate)[/dim]")
    else:
        console.print("  [dim]No Excel workbook loaded[/dim]")
        console.print("  [dim](Use 'use excel <file>' to load)[/dim]")

    # PowerPoint context
    console.print("\n[bold magenta]PowerPoint Context:[/bold magenta]")
    if ctx.ppt_presentation_path:
        console.print(f"  Presentation: {ctx.ppt_presentation_name}")
        console.print(f"  Path: {ctx.ppt_presentation_path}")
        console.print(f"  Active Slide: {ctx.ppt_slide_number or 'None'}")
        if ctx.mode != "ppt":
            console.print("  [dim](Use 'switch ppt' to activate)[/dim]")
    else:
        console.print("  [dim]No PowerPoint presentation loaded[/dim]")
        console.print("  [dim](Use 'use ppt <file>' to load)[/dim]")

    console.print()


def handle_use_command(ctx: UnifiedShellContext, args: List[str]) -> bool:
    """
    Handle 'use' command for mode switching with file loading

    Examples:
        use excel "sales.xlsx"
        use ppt "report.pptx"
    """
    if len(args) < 2:
        console.print("[red]Usage: use excel <file> | use ppt <file>[/red]")
        return True

    mode = args[0].lower()
    file_path = args[1]

    if mode == "excel":
        # Switch to Excel mode and open file
        from pyhub_office_automation.cli.main import app as main_app

        result = runner.invoke(main_app, ["excel", "workbook-open", "--file-path", file_path])

        if result.exit_code == 0:
            ctx.mode = "excel"
            ctx.excel_workbook_path = str(Path(file_path).absolute())
            ctx.excel_workbook_name = Path(file_path).name

            # Get active sheet (would need actual Excel query here)
            # For now, assume first sheet is active
            console.print(f"[green]✓ Excel workbook: {ctx.excel_workbook_name}[/green]")
            console.print("[green]✓ Mode switched to Excel[/green]")
        else:
            console.print(f"[red]Failed to open Excel file: {file_path}[/red]")
            console.print(result.stdout)

    elif mode == "ppt":
        # Switch to PowerPoint mode and open file
        from pyhub_office_automation.cli.main import app as main_app

        result = runner.invoke(main_app, ["ppt", "presentation-open", "--file-path", file_path])

        if result.exit_code == 0:
            ctx.mode = "ppt"
            ctx.ppt_presentation_path = str(Path(file_path).absolute())
            ctx.ppt_presentation_name = Path(file_path).name
            ctx.ppt_slide_number = 1  # Default to first slide

            console.print(f"[magenta]✓ PowerPoint presentation: {ctx.ppt_presentation_name}[/magenta]")
            console.print("[magenta]✓ Mode switched to PowerPoint[/magenta]")
        else:
            console.print(f"[red]Failed to open PowerPoint file: {file_path}[/red]")
            console.print(result.stdout)

    else:
        console.print(f"[red]Unknown mode: {mode}. Use 'excel' or 'ppt'.[/red]")

    return True


def handle_switch_command(ctx: UnifiedShellContext, args: List[str]) -> bool:
    """
    Handle 'switch' command for mode switching without file loading

    Examples:
        switch excel
        switch ppt
    """
    if len(args) < 1:
        console.print("[red]Usage: switch excel | switch ppt[/red]")
        return True

    mode = args[0].lower()

    if mode == "excel":
        if ctx.excel_workbook_path:
            ctx.mode = "excel"
            console.print(f"[green]✓ Switched to Excel mode: {ctx.excel_workbook_name}[/green]")
        else:
            console.print("[yellow]No Excel workbook loaded. Use 'use excel <file>' first.[/yellow]")

    elif mode == "ppt":
        if ctx.ppt_presentation_path:
            ctx.mode = "ppt"
            console.print(f"[magenta]✓ Switched to PowerPoint mode: {ctx.ppt_presentation_name}[/magenta]")
        else:
            console.print("[yellow]No PowerPoint presentation loaded. Use 'use ppt <file>' first.[/yellow]")

    else:
        console.print(f"[red]Unknown mode: {mode}. Use 'excel' or 'ppt'.[/red]")

    return True


def execute_excel_command(ctx: UnifiedShellContext, command: str, args: List[str]):
    """Execute Excel command with context injection"""
    from pyhub_office_automation.cli.main import app as main_app

    # Build command arguments
    cmd_args = ["excel", command]

    # Context injection
    has_file_arg = any(arg in ["--file-path", "--workbook-name"] for arg in args)
    has_sheet_arg = "--sheet" in args

    if not has_file_arg and ctx.excel_workbook_path:
        cmd_args.extend(["--file-path", ctx.excel_workbook_path])

    if not has_sheet_arg and ctx.excel_sheet:
        cmd_args.extend(["--sheet", ctx.excel_sheet])

    cmd_args.extend(args)

    # Execute command
    result = runner.invoke(main_app, cmd_args)
    console.print(result.stdout)

    # Update context based on command
    if command == "use" and len(args) >= 2 and args[0] == "sheet":
        ctx.excel_sheet = args[1]


def execute_ppt_command(ctx: UnifiedShellContext, command: str, args: List[str]):
    """Execute PowerPoint command with context injection"""
    from pyhub_office_automation.cli.main import app as main_app

    # Build command arguments
    cmd_args = ["ppt", command]

    # Context injection
    has_file_arg = "--file-path" in args
    has_slide_arg = "--slide-number" in args

    if not has_file_arg and ctx.ppt_presentation_path:
        cmd_args.extend(["--file-path", ctx.ppt_presentation_path])

    # Slide-specific commands
    slide_commands = [
        "content-add-text",
        "content-add-image",
        "content-add-shape",
        "content-add-table",
        "content-add-chart",
        "content-add-video",
        "content-add-audio",
        "slide-delete",
        "slide-duplicate",
        "layout-apply",
    ]

    if command in slide_commands and not has_slide_arg and ctx.ppt_slide_number:
        cmd_args.extend(["--slide-number", str(ctx.ppt_slide_number)])

    cmd_args.extend(args)

    # Execute command
    result = runner.invoke(main_app, cmd_args)
    console.print(result.stdout)

    # Update context based on command
    if command == "use" and len(args) >= 2 and args[0] == "slide":
        try:
            ctx.ppt_slide_number = int(args[1])
        except ValueError:
            pass


def execute_unified_command(ctx: UnifiedShellContext, command: str, args: List[str]) -> bool:
    """
    Execute command based on current mode

    Returns:
        True to continue shell, False to exit
    """
    # Unified shell commands (available in all modes)
    if command == "help":
        show_unified_help(ctx)
        return True

    if command == "show":
        if args and args[0] == "context":
            show_unified_context(ctx)
        else:
            console.print("[yellow]Usage: show context[/yellow]")
        return True

    if command == "clear":
        console.clear()
        return True

    if command in ["exit", "quit"]:
        console.print("[cyan]Goodbye from Unified Shell![/cyan]")
        return False

    # Mode switching commands
    if command == "use":
        handle_use_command(ctx, args)
        return True

    if command == "switch":
        handle_switch_command(ctx, args)
        return True

    # Mode-specific commands
    if ctx.mode == "excel":
        if command in EXCEL_COMMANDS or command == "use":
            execute_excel_command(ctx, command, args)
            return True

    elif ctx.mode == "ppt":
        if command in PPT_COMMANDS or command == "use":
            execute_ppt_command(ctx, command, args)
            return True

    # No mode active or unknown command
    if ctx.mode == "none":
        console.print("[yellow]No active mode. Use 'use excel <file>' or 'use ppt <file>' first.[/yellow]")
    else:
        console.print(f"[red]Unknown command: {command}. Type 'help' for available commands.[/red]")

    return True


def unified_shell():
    """
    Start unified shell mode for Excel and PowerPoint

    Provides a single REPL interface for managing both Excel workbooks
    and PowerPoint presentations with seamless context switching.
    """
    console.print("[bold cyan]Unified Shell Mode (Excel + PowerPoint)[/bold cyan]")
    console.print("Type 'help' for available commands, 'exit' to quit.\n")

    # Initialize context
    ctx = UnifiedShellContext()

    # Setup prompt session with history
    history_file = Path.home() / ".oa_unified_shell_history"
    session: PromptSession = PromptSession(history=FileHistory(str(history_file)), completer=UnifiedShellCompleter(ctx))

    # Main REPL loop
    while True:
        try:
            # Get user input with dynamic prompt
            user_input = session.prompt(ctx.get_prompt_text())

            if not user_input.strip():
                continue

            # Parse command
            try:
                tokens = shlex.split(user_input)
            except ValueError as e:
                console.print(f"[red]Parse error: {e}[/red]")
                continue

            if not tokens:
                continue

            command = tokens[0]
            args = tokens[1:]

            # Execute command
            should_continue = execute_unified_command(ctx, command, args)
            if not should_continue:
                break

        except KeyboardInterrupt:
            console.print("\n[yellow]Use 'exit' or 'quit' to leave the shell.[/yellow]")
            continue

        except EOFError:
            console.print("\n[cyan]Goodbye from Unified Shell![/cyan]")
            break

        except Exception as e:
            console.print(f"[red]Error: {e}[/red]")
            import traceback

            traceback.print_exc()
            continue
