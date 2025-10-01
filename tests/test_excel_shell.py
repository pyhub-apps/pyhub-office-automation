"""
Test cases for Excel Shell Mode
"""

from unittest.mock import MagicMock, Mock, patch

import pytest

from pyhub_office_automation.shell.excel_shell import (
    ExcelShellCompleter,
    ExcelShellContext,
    execute_shell_command,
    parse_shell_args,
)


class TestExcelShellContext:
    """Test ExcelShellContext class"""

    def test_context_initialization(self):
        """Test context initialization"""
        ctx = ExcelShellContext()
        assert ctx.workbook_name is None
        assert ctx.sheet_name is None
        assert ctx.app is None
        assert ctx.workbook is None

    def test_get_prompt_text_no_context(self):
        """Test prompt text with no context"""
        ctx = ExcelShellContext()
        prompt = ctx.get_prompt_text()
        assert prompt == "[Excel: None > None] > "

    def test_get_prompt_text_with_context(self):
        """Test prompt text with context"""
        ctx = ExcelShellContext()
        ctx.workbook_name = "test.xlsx"
        ctx.sheet_name = "Sheet1"
        prompt = ctx.get_prompt_text()
        assert prompt == "[Excel: test.xlsx > Sheet1] > "

    @patch("pyhub_office_automation.shell.excel_shell.get_workbook")
    @patch("pyhub_office_automation.shell.excel_shell.console")
    def test_update_workbook(self, mock_console, mock_get_workbook):
        """Test workbook update"""
        # Create mock workbook
        mock_wb = MagicMock()
        mock_wb.name = "test.xlsx"
        mock_sheet = MagicMock()
        mock_sheet.name = "Sheet1"

        # Create mock sheets collection
        mock_sheets = MagicMock()
        mock_sheets.__len__ = Mock(return_value=1)
        mock_sheets.active = mock_sheet
        mock_wb.sheets = mock_sheets

        mock_get_workbook.return_value = mock_wb

        ctx = ExcelShellContext()
        ctx.update_workbook("test.xlsx")

        assert ctx.workbook_name == "test.xlsx"
        assert ctx.sheet_name == "Sheet1"
        mock_get_workbook.assert_called_once_with(workbook_name="test.xlsx")

    @patch("pyhub_office_automation.shell.excel_shell.console")
    def test_update_sheet_no_workbook(self, mock_console):
        """Test sheet update without workbook"""
        ctx = ExcelShellContext()
        ctx.update_sheet("Sheet1")

        # Should print error message
        assert mock_console.print.called

    @patch("pyhub_office_automation.shell.excel_shell.console")
    def test_update_sheet_with_workbook(self, mock_console):
        """Test sheet update with workbook"""
        # Create mock workbook
        mock_wb = MagicMock()
        mock_sheet = MagicMock()
        mock_sheet.name = "Sheet1"
        mock_wb.sheets = {"Sheet1": mock_sheet}

        ctx = ExcelShellContext()
        ctx.workbook = mock_wb
        ctx.workbook_name = "test.xlsx"
        ctx.update_sheet("Sheet1")

        assert ctx.sheet_name == "Sheet1"
        mock_sheet.activate.assert_called_once()


class TestParseShellArgs:
    """Test parse_shell_args function"""

    def test_simple_command(self):
        """Test simple command parsing"""
        result = parse_shell_args("help")
        assert result == ["help"]

    def test_command_with_args(self):
        """Test command with arguments"""
        result = parse_shell_args("use sheet TestData")
        assert result == ["use", "sheet", "TestData"]

    def test_quoted_arguments(self):
        """Test quoted arguments"""
        result = parse_shell_args('range-read --range "A1:C10"')
        assert result == ["range-read", "--range", "A1:C10"]

    def test_mixed_arguments(self):
        """Test mixed quoted and unquoted"""
        result = parse_shell_args("range-write --range A1 --data '[[1,2,3]]'")
        assert result == ["range-write", "--range", "A1", "--data", "[[1,2,3]]"]

    def test_empty_string(self):
        """Test empty string"""
        result = parse_shell_args("")
        assert result == []


class TestExecuteShellCommand:
    """Test execute_shell_command function"""

    @patch("pyhub_office_automation.shell.excel_shell.console")
    def test_exit_command(self, mock_console):
        """Test exit command"""
        ctx = ExcelShellContext()
        result = execute_shell_command(ctx, "exit")
        assert result is False  # Should return False to exit loop

    @patch("pyhub_office_automation.shell.excel_shell.console")
    def test_quit_command(self, mock_console):
        """Test quit command"""
        ctx = ExcelShellContext()
        result = execute_shell_command(ctx, "quit")
        assert result is False

    @patch("pyhub_office_automation.shell.excel_shell.console")
    def test_clear_command(self, mock_console):
        """Test clear command"""
        ctx = ExcelShellContext()
        result = execute_shell_command(ctx, "clear")
        assert result is True
        mock_console.clear.assert_called_once()

    @patch("pyhub_office_automation.shell.excel_shell.show_help")
    def test_help_command(self, mock_show_help):
        """Test help command"""
        ctx = ExcelShellContext()
        result = execute_shell_command(ctx, "help")
        assert result is True
        mock_show_help.assert_called_once()

    @patch("pyhub_office_automation.shell.excel_shell.show_context")
    def test_show_context_command(self, mock_show_context):
        """Test show context command"""
        ctx = ExcelShellContext()
        result = execute_shell_command(ctx, "show context")
        assert result is True
        mock_show_context.assert_called_once_with(ctx)

    @patch("pyhub_office_automation.shell.excel_shell.console")
    def test_unknown_command(self, mock_console):
        """Test unknown command"""
        ctx = ExcelShellContext()
        result = execute_shell_command(ctx, "unknown_command")
        assert result is True  # Should continue
        # Should print error message
        assert mock_console.print.called


class TestExcelShellCompleter:
    """Test ExcelShellCompleter class"""

    def test_completer_initialization(self):
        """Test completer initialization"""
        ctx = ExcelShellContext()
        completer = ExcelShellCompleter(ctx)
        assert completer.context == ctx
        assert len(completer.shell_commands) > 0

    def test_command_completion(self):
        """Test command name completion"""
        ctx = ExcelShellContext()
        completer = ExcelShellCompleter(ctx)

        # Create mock document
        mock_doc = Mock()
        mock_doc.text_before_cursor = "ran"

        completions = list(completer.get_completions(mock_doc, None))
        # Should suggest range-read and range-write
        completion_texts = [c.text for c in completions]
        assert "range-read" in completion_texts
        assert "range-write" in completion_texts

    def test_use_subcommand_completion(self):
        """Test 'use' subcommand completion"""
        ctx = ExcelShellContext()
        completer = ExcelShellCompleter(ctx)

        # Create mock document
        mock_doc = Mock()
        mock_doc.text_before_cursor = "use w"

        completions = list(completer.get_completions(mock_doc, None))
        completion_texts = [c.text for c in completions]
        assert "workbook" in completion_texts


# Integration test placeholder
class TestShellIntegration:
    """Integration tests (require actual Excel)"""

    @pytest.mark.skip(reason="Requires manual Excel interaction")
    def test_full_workflow(self):
        """Test full shell workflow"""
        # This would require actual Excel instance
        pass


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
