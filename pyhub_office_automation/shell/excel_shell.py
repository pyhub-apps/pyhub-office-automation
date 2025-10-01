"""
Interactive Excel shell for stateful Excel automation
Issue #85: https://github.com/pyhub-apps/pyhub-office-automation/issues/85
"""

import json
import shlex
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import typer
import xlwings as xw
from prompt_toolkit import PromptSession
from prompt_toolkit.completion import Completer, Completion
from prompt_toolkit.history import FileHistory
from prompt_toolkit.styles import Style
from rich.console import Console
from rich.table import Table

# Chart Commands (7)
from pyhub_office_automation.excel.chart_add import chart_add
from pyhub_office_automation.excel.chart_configure import chart_configure
from pyhub_office_automation.excel.chart_delete import chart_delete
from pyhub_office_automation.excel.chart_export import chart_export
from pyhub_office_automation.excel.chart_list import chart_list
from pyhub_office_automation.excel.chart_pivot_create import chart_pivot_create
from pyhub_office_automation.excel.chart_position import chart_position

# Data Commands (2)
from pyhub_office_automation.excel.data_analyze import data_analyze
from pyhub_office_automation.excel.data_transform import data_transform

# Workbook Commands (5)
from pyhub_office_automation.excel.metadata_generate import metadata_generate

# Pivot Commands (5)
from pyhub_office_automation.excel.pivot_configure import pivot_configure
from pyhub_office_automation.excel.pivot_create import pivot_create
from pyhub_office_automation.excel.pivot_delete import pivot_delete
from pyhub_office_automation.excel.pivot_list import pivot_list
from pyhub_office_automation.excel.pivot_refresh import pivot_refresh

# Excel 명령어 함수 import - ALL COMMANDS
# Range Commands (3)
from pyhub_office_automation.excel.range_convert import range_convert
from pyhub_office_automation.excel.range_read import range_read
from pyhub_office_automation.excel.range_write import range_write

# Shape Commands (6)
from pyhub_office_automation.excel.shape_add import shape_add
from pyhub_office_automation.excel.shape_delete import shape_delete
from pyhub_office_automation.excel.shape_format import shape_format
from pyhub_office_automation.excel.shape_group import shape_group
from pyhub_office_automation.excel.shape_list import shape_list

# Sheet Commands (4)
from pyhub_office_automation.excel.sheet_activate import sheet_activate
from pyhub_office_automation.excel.sheet_add import sheet_add
from pyhub_office_automation.excel.sheet_delete import sheet_delete
from pyhub_office_automation.excel.sheet_rename import sheet_rename

# Slicer Commands (4)
from pyhub_office_automation.excel.slicer_add import slicer_add
from pyhub_office_automation.excel.slicer_connect import slicer_connect
from pyhub_office_automation.excel.slicer_list import slicer_list
from pyhub_office_automation.excel.slicer_position import slicer_position

# Table Commands (8)
from pyhub_office_automation.excel.table_analyze import table_analyze
from pyhub_office_automation.excel.table_create import table_create
from pyhub_office_automation.excel.table_list import table_list
from pyhub_office_automation.excel.table_read import table_read
from pyhub_office_automation.excel.table_sort import table_sort
from pyhub_office_automation.excel.table_sort_clear import table_sort_clear
from pyhub_office_automation.excel.table_sort_info import table_sort_info
from pyhub_office_automation.excel.table_write import table_write
from pyhub_office_automation.excel.textbox_add import textbox_add

# Utility
from pyhub_office_automation.excel.utils import get_active_workbook, get_workbook, normalize_path
from pyhub_office_automation.excel.workbook_create import workbook_create
from pyhub_office_automation.excel.workbook_info import workbook_info
from pyhub_office_automation.excel.workbook_list import workbook_list
from pyhub_office_automation.excel.workbook_open import workbook_open

console = Console()


@dataclass
class ExcelShellContext:
    """Excel shell session context"""

    workbook_name: Optional[str] = None
    sheet_name: Optional[str] = None
    app: Optional[xw.App] = None
    workbook: Optional[xw.Book] = None

    def get_prompt_text(self) -> str:
        """Generate prompt text based on current context"""
        wb = self.workbook_name or "None"
        sheet = self.sheet_name or "None"
        return f"[Excel: {wb} > {sheet}] > "

    def update_workbook(self, workbook_name: Optional[str] = None):
        """Update workbook context"""
        if workbook_name:
            try:
                self.workbook = get_workbook(workbook_name=workbook_name)
                self.workbook_name = workbook_name
                if self.workbook and len(self.workbook.sheets) > 0:
                    self.sheet_name = self.workbook.sheets.active.name
                console.print(f"[green]Switched to workbook: {workbook_name}[/green]")
                console.print(f"[dim]Active sheet: {self.sheet_name}[/dim]")
            except Exception as e:
                console.print(f"[red]Failed to switch workbook: {e}[/red]")
        else:
            # Use active workbook
            try:
                self.workbook = get_active_workbook()
                if self.workbook:
                    self.workbook_name = self.workbook.name
                    if len(self.workbook.sheets) > 0:
                        self.sheet_name = self.workbook.sheets.active.name
                    console.print(f"[green]Using active workbook: {self.workbook_name}[/green]")
                    console.print(f"[dim]Active sheet: {self.sheet_name}[/dim]")
            except Exception as e:
                console.print(f"[red]No active workbook found: {e}[/red]")

    def update_sheet(self, sheet_name: str):
        """Update sheet context"""
        if not self.workbook:
            console.print("[red]No workbook selected[/red]")
            console.print("[yellow]Use 'use workbook <name>' first[/yellow]")
            return

        try:
            sheet = self.workbook.sheets[sheet_name]
            sheet.activate()
            self.sheet_name = sheet_name
            console.print(f"[green]Switched to sheet: {sheet_name}[/green]")
        except Exception as e:
            console.print(f"[red]Failed to switch sheet: {e}[/red]")
            console.print("[yellow]Tip: Use 'sheets' to see available sheets[/yellow]")


class ExcelShellCompleter(Completer):
    """Autocomplete for Excel shell commands"""

    def __init__(self, context: ExcelShellContext):
        self.context = context
        # Shell commands
        self.shell_commands = [
            "use",
            "show",
            "workbooks",
            "sheets",
            "clear",
            "help",
            "exit",
            "quit",
        ]
        # All Excel commands (44)
        self.excel_commands = [
            # Range Commands
            "range-read",
            "range-write",
            "range-convert",
            # Workbook Commands
            "workbook-list",
            "workbook-open",
            "workbook-create",
            "workbook-info",
            "metadata-generate",
            # Sheet Commands
            "sheet-activate",
            "sheet-add",
            "sheet-delete",
            "sheet-rename",
            # Table Commands
            "table-create",
            "table-list",
            "table-read",
            "table-sort",
            "table-sort-clear",
            "table-sort-info",
            "table-write",
            "table-analyze",
            # Data Commands
            "data-analyze",
            "data-transform",
            # Chart Commands
            "chart-add",
            "chart-configure",
            "chart-delete",
            "chart-export",
            "chart-list",
            "chart-pivot-create",
            "chart-position",
            # Pivot Commands
            "pivot-configure",
            "pivot-create",
            "pivot-delete",
            "pivot-list",
            "pivot-refresh",
            # Shape Commands
            "shape-add",
            "shape-delete",
            "shape-format",
            "shape-group",
            "shape-list",
            "textbox-add",
            # Slicer Commands
            "slicer-add",
            "slicer-connect",
            "slicer-list",
            "slicer-position",
        ]
        self.all_commands = self.shell_commands + self.excel_commands

    def get_completions(self, document, complete_event):
        """Get completions for current input"""
        text = document.text_before_cursor
        words = text.split()

        # First word completion (commands)
        if len(words) <= 1:
            for cmd in self.all_commands:
                if cmd.startswith(text.lower()):
                    yield Completion(cmd, start_position=-len(text))

        # "use" subcommand completion
        elif len(words) == 2 and words[0] == "use":
            subcommands = ["workbook", "sheet"]
            for sub in subcommands:
                if sub.startswith(words[1].lower()):
                    yield Completion(sub, start_position=-len(words[1]))

        # "use workbook" - list open workbooks
        elif len(words) == 3 and words[0] == "use" and words[1] == "workbook":
            try:
                if sys.platform == "win32":
                    app = xw.apps.active
                    if app:
                        for book in app.books:
                            if book.name.lower().startswith(words[2].lower()):
                                yield Completion(book.name, start_position=-len(words[2]))
            except Exception:
                pass

        # "use sheet" - list sheets in current workbook
        elif len(words) == 3 and words[0] == "use" and words[1] == "sheet":
            if self.context.workbook:
                try:
                    for sheet in self.context.workbook.sheets:
                        if sheet.name.lower().startswith(words[2].lower()):
                            yield Completion(sheet.name, start_position=-len(words[2]))
                except Exception:
                    pass

        # "show" subcommand completion
        elif len(words) == 2 and words[0] == "show":
            subcommands = ["context", "workbooks", "sheets"]
            for sub in subcommands:
                if sub.startswith(words[1].lower()):
                    yield Completion(sub, start_position=-len(words[1]))


def parse_shell_args(command_str: str):
    """Parse shell command string into arguments list"""
    try:
        # Use shlex to properly handle quoted arguments
        return shlex.split(command_str)
    except ValueError as e:
        console.print(f"[red]Argument parsing error: {e}[/red]")
        return None


def execute_excel_command(ctx: ExcelShellContext, cmd: str, args: list) -> bool:
    """
    Execute Excel commands with context injection

    Args:
        ctx: Current shell context
        cmd: Command name (e.g., "range-read")
        args: Command arguments list

    Returns:
        True if should continue shell, False if should exit
    """
    # Excel 명령어 매핑 (ALL 44 commands)
    excel_commands = {
        # Range Commands (3)
        "range-read": range_read,
        "range-write": range_write,
        "range-convert": range_convert,
        # Workbook Commands (5)
        "workbook-list": workbook_list,
        "workbook-open": workbook_open,
        "workbook-create": workbook_create,
        "workbook-info": workbook_info,
        "metadata-generate": metadata_generate,
        # Sheet Commands (4)
        "sheet-activate": sheet_activate,
        "sheet-add": sheet_add,
        "sheet-delete": sheet_delete,
        "sheet-rename": sheet_rename,
        # Table Commands (8)
        "table-create": table_create,
        "table-list": table_list,
        "table-read": table_read,
        "table-sort": table_sort,
        "table-sort-clear": table_sort_clear,
        "table-sort-info": table_sort_info,
        "table-write": table_write,
        "table-analyze": table_analyze,
        # Data Commands (2)
        "data-analyze": data_analyze,
        "data-transform": data_transform,
        # Chart Commands (7)
        "chart-add": chart_add,
        "chart-configure": chart_configure,
        "chart-delete": chart_delete,
        "chart-export": chart_export,
        "chart-list": chart_list,
        "chart-pivot-create": chart_pivot_create,
        "chart-position": chart_position,
        # Pivot Commands (5)
        "pivot-configure": pivot_configure,
        "pivot-create": pivot_create,
        "pivot-delete": pivot_delete,
        "pivot-list": pivot_list,
        "pivot-refresh": pivot_refresh,
        # Shape Commands (6)
        "shape-add": shape_add,
        "shape-delete": shape_delete,
        "shape-format": shape_format,
        "shape-group": shape_group,
        "shape-list": shape_list,
        "textbox-add": textbox_add,
        # Slicer Commands (4)
        "slicer-add": slicer_add,
        "slicer-connect": slicer_connect,
        "slicer-list": slicer_list,
        "slicer-position": slicer_position,
    }

    if cmd not in excel_commands:
        return False  # Not an Excel command

    command_func = excel_commands[cmd]

    try:
        # Context injection: add workbook and sheet if not specified
        injected_args = []
        has_workbook_arg = False
        has_sheet_arg = False

        # Check if arguments already contain workbook/sheet options
        i = 0
        while i < len(args):
            arg = args[i]
            if arg in ["--workbook-name", "--file-path"]:
                has_workbook_arg = True
                injected_args.append(arg)
                if i + 1 < len(args):
                    injected_args.append(args[i + 1])
                    i += 2
                else:
                    i += 1
            elif arg == "--sheet":
                has_sheet_arg = True
                injected_args.append(arg)
                if i + 1 < len(args):
                    injected_args.append(args[i + 1])
                    i += 2
                else:
                    i += 1
            else:
                injected_args.append(arg)
                i += 1

        # Inject context if not specified
        if not has_workbook_arg and ctx.workbook_name:
            injected_args.extend(["--workbook-name", ctx.workbook_name])

        if not has_sheet_arg and ctx.sheet_name:
            injected_args.extend(["--sheet", ctx.sheet_name])

        # Create a Typer context and invoke the command
        import typer as typer_module
        from typer.testing import CliRunner

        # Create a temporary app for command execution
        temp_app = typer_module.Typer()
        temp_app.command()(command_func)

        runner = CliRunner(mix_stderr=False)
        result = runner.invoke(temp_app, injected_args)

        # Display output
        if result.stdout:
            console.print(result.stdout, end="")
        if result.stderr:
            console.print(f"[red]{result.stderr}[/red]", end="")

        return True

    except Exception as e:
        console.print(f"[red]Error executing {cmd}: {e}[/red]")
        import traceback

        console.print(f"[dim]{traceback.format_exc()}[/dim]")
        return True


def execute_shell_command(ctx: ExcelShellContext, command: str) -> bool:
    """
    Execute shell-specific commands
    Returns True if should continue, False if should exit
    """
    # Parse command arguments
    parts = parse_shell_args(command.strip())
    if not parts:
        return True

    cmd = parts[0].lower()

    # Exit commands
    if cmd in ["exit", "quit"]:
        console.print("Shell session ended", style="green")
        return False

    # Clear command
    elif cmd == "clear":
        console.clear()
        return True

    # Help command
    elif cmd == "help":
        show_help()
        return True

    # Show commands
    elif cmd == "show":
        if len(parts) < 2:
            console.print("Usage: show [context|workbooks|sheets]", style="yellow")
        elif parts[1] == "context":
            show_context(ctx)
        elif parts[1] == "workbooks":
            show_workbooks()
        elif parts[1] == "sheets":
            show_sheets(ctx)
        return True

    # Use commands
    elif cmd == "use":
        if len(parts) < 3:
            console.print("Usage: use [workbook|sheet] <name>", style="yellow")
        elif parts[1] == "workbook":
            ctx.update_workbook(parts[2])
        elif parts[1] == "sheet":
            ctx.update_sheet(parts[2])
        return True

    # Shortcut commands
    elif cmd == "workbooks":
        show_workbooks()
        return True

    elif cmd == "sheets":
        show_sheets(ctx)
        return True

    # Excel commands - delegate to actual command implementation (44 commands)
    elif cmd in [
        # Range
        "range-read",
        "range-write",
        "range-convert",
        # Workbook
        "workbook-list",
        "workbook-open",
        "workbook-create",
        "workbook-info",
        "metadata-generate",
        # Sheet
        "sheet-activate",
        "sheet-add",
        "sheet-delete",
        "sheet-rename",
        # Table
        "table-create",
        "table-list",
        "table-read",
        "table-sort",
        "table-sort-clear",
        "table-sort-info",
        "table-write",
        "table-analyze",
        # Data
        "data-analyze",
        "data-transform",
        # Chart
        "chart-add",
        "chart-configure",
        "chart-delete",
        "chart-export",
        "chart-list",
        "chart-pivot-create",
        "chart-position",
        # Pivot
        "pivot-configure",
        "pivot-create",
        "pivot-delete",
        "pivot-list",
        "pivot-refresh",
        # Shape
        "shape-add",
        "shape-delete",
        "shape-format",
        "shape-group",
        "shape-list",
        "textbox-add",
        # Slicer
        "slicer-add",
        "slicer-connect",
        "slicer-list",
        "slicer-position",
    ]:
        return execute_excel_command(ctx, cmd, parts[1:])

    else:
        console.print(f"Unknown command: {cmd}", style="red")
        console.print("Type 'help' for available commands", style="yellow")
        return True


def show_help():
    """Show help information"""
    # Shell-specific commands
    table1 = Table(title="Shell Commands")
    table1.add_column("Command", style="cyan")
    table1.add_column("Description", style="white")

    table1.add_row("use workbook <name>", "Switch to specified workbook")
    table1.add_row("use sheet <name>", "Switch to specified sheet")
    table1.add_row("show context", "Show current context (workbook, sheet)")
    table1.add_row("show workbooks", "List all open workbooks")
    table1.add_row("show sheets", "List sheets in current workbook")
    table1.add_row("workbooks", "Shortcut for 'show workbooks'")
    table1.add_row("sheets", "Shortcut for 'show sheets'")
    table1.add_row("clear", "Clear screen")
    table1.add_row("help", "Show this help message")
    table1.add_row("exit / quit", "Exit shell mode")

    console.print(table1)
    console.print()

    # Excel commands
    table2 = Table(title="Excel Commands (Context Auto-Injected) - 44 Commands Available")
    table2.add_column("Category", style="magenta")
    table2.add_column("Examples", style="white")

    table2.add_row("Range (3)", "range-read --range A1:C10, range-write, range-convert")
    table2.add_row("Workbook (5)", "workbook-list, workbook-info, workbook-create, metadata-generate")
    table2.add_row("Sheet (4)", "sheet-add --name NewSheet, sheet-delete, sheet-activate")
    table2.add_row("Table (8)", "table-list, table-read, table-create, table-sort")
    table2.add_row("Data (2)", "data-analyze, data-transform")
    table2.add_row("Chart (7)", "chart-add, chart-list, chart-configure, chart-export")
    table2.add_row("Pivot (5)", "pivot-create, pivot-list, pivot-refresh, pivot-configure")
    table2.add_row("Shape (6)", "shape-add, shape-list, textbox-add, shape-group")
    table2.add_row("Slicer (4)", "slicer-add, slicer-list, slicer-connect, slicer-position")

    console.print(table2)
    console.print("\n[yellow]Note: Context (workbook/sheet) is automatically injected![/yellow]")
    console.print("[dim]Tip: Use Tab for command autocomplete, <command> --help for details[/dim]")


def show_context(ctx: ExcelShellContext):
    """Show current context"""
    table = Table(title="Current Context")
    table.add_column("Property", style="cyan")
    table.add_column("Value", style="white")

    table.add_row("Workbook", ctx.workbook_name or "None")
    table.add_row("Sheet", ctx.sheet_name or "None")
    table.add_row("Active", "Yes" if ctx.workbook else "No")

    console.print(table)


def show_workbooks():
    """Show all open workbooks"""
    try:
        if sys.platform != "win32":
            console.print("[yellow]Workbook listing only available on Windows[/yellow]")
            return

        app = xw.apps.active
        if not app:
            console.print("[yellow]No active Excel application[/yellow]")
            return

        table = Table(title="Open Workbooks")
        table.add_column("#", style="cyan")
        table.add_column("Name", style="white")
        table.add_column("Path", style="dim")

        for i, book in enumerate(app.books, 1):
            path = book.fullname if hasattr(book, "fullname") else "Unsaved"
            table.add_row(str(i), book.name, path)

        console.print(table)

    except Exception as e:
        console.print(f"[red]Error listing workbooks: {e}[/red]")


def show_sheets(ctx: ExcelShellContext):
    """Show sheets in current workbook"""
    if not ctx.workbook:
        console.print("[yellow]No workbook selected[/yellow]")
        return

    try:
        table = Table(title=f"Sheets in {ctx.workbook_name}")
        table.add_column("#", style="cyan")
        table.add_column("Name", style="white")
        table.add_column("Active", style="green")

        for i, sheet in enumerate(ctx.workbook.sheets, 1):
            is_active = "✓" if sheet.name == ctx.sheet_name else ""
            table.add_row(str(i), sheet.name, is_active)

        console.print(table)

    except Exception as e:
        console.print(f"[red]Error listing sheets: {e}[/red]")


def excel_shell(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="Excel file path to open"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help="Name of already open workbook"),
):
    """
    Start interactive Excel shell mode

    Stateful REPL interface for Excel automation with context management,
    autocomplete, and command history.

    Examples:
        oa excel shell
        oa excel shell --file-path "report.xlsx"
        oa excel shell --workbook-name "Book1.xlsx"
    """
    console.print("\n[bold cyan]Excel Interactive Shell[/bold cyan]")
    console.print("[dim]Type 'help' for available commands, 'exit' to quit[/dim]")
    console.print("[dim]Tab for autocomplete, Up/Down for history[/dim]\n")

    # Initialize context
    ctx = ExcelShellContext()

    # Open or connect to workbook
    if file_path:
        try:
            file_path = normalize_path(file_path)
            if not Path(file_path).exists():
                console.print(f"[red]File not found: {file_path}[/red]")
                raise typer.Exit(1)

            workbook = xw.Book(file_path)
            ctx.workbook = workbook
            ctx.workbook_name = workbook.name
            if len(workbook.sheets) > 0:
                ctx.sheet_name = workbook.sheets.active.name
            console.print(f"[green]Opened workbook: {ctx.workbook_name}[/green]")
            if ctx.sheet_name:
                console.print(f"[dim]Active sheet: {ctx.sheet_name}[/dim]")
            console.print(f"[dim]Total sheets: {len(workbook.sheets)}[/dim]")

        except Exception as e:
            console.print(f"[red]Failed to open workbook: {e}[/red]")
            raise typer.Exit(1)

    elif workbook_name:
        ctx.update_workbook(workbook_name)

    else:
        # Use active workbook
        ctx.update_workbook()

    # Setup prompt session
    history_file = Path.home() / ".oa_excel_shell_history"
    session = PromptSession(
        history=FileHistory(str(history_file)),
        completer=ExcelShellCompleter(ctx),
        complete_while_typing=True,
    )

    # REPL loop
    while True:
        try:
            prompt_text = ctx.get_prompt_text()
            user_input = session.prompt(prompt_text)

            if not user_input.strip():
                continue

            # Execute command
            should_continue = execute_shell_command(ctx, user_input)
            if not should_continue:
                break

        except KeyboardInterrupt:
            console.print("\n[yellow]Use 'exit' or 'quit' to leave shell[/yellow]")
            continue

        except EOFError:
            console.print("\n[green]Shell session ended[/green]")
            break

        except Exception as e:
            console.print(f"[red]Error: {e}[/red]")
            continue
