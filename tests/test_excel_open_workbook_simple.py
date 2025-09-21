"""
Excel workbook-open 명령어 간단 테스트
"""

import json
from pathlib import Path
from unittest.mock import Mock, patch

import pytest
from typer.testing import CliRunner

from pyhub_office_automation.cli.main import excel_app


class TestWorkbookOpenSimple:
    """Excel workbook-open 명령어 간단 테스트 클래스"""

    def test_help_option(self):
        """도움말 옵션 테스트"""
        runner = CliRunner()

        result = runner.invoke(excel_app, ["workbook-open", "--help"])

        assert result.exit_code == 0
        assert "Excel 워크북을 열거나 기존 워크북의 정보를 가져옵니다" in result.output
        assert "--file-path" in result.output
        assert "--visible" in result.output
        assert "--format" in result.output

    def test_version_option(self):
        """버전 옵션 테스트 - Typer에서는 개별 명령어에 --version이 없음"""
        runner = CliRunner()

        # Typer에서는 개별 명령어에 --version이 없으므로 실패해야 함
        result = runner.invoke(excel_app, ["workbook-open", "--version"])

        assert result.exit_code != 0  # 실패해야 함
        assert "No such option" in result.output or "Unrecognized arguments" in result.output

    def test_file_not_found_error(self):
        """파일이 존재하지 않는 경우 테스트"""
        runner = CliRunner()
        non_existent_file = "/path/to/non_existent_file.xlsx"

        result = runner.invoke(excel_app, ["workbook-open", "--file-path", non_existent_file, "--format", "json"])

        assert result.exit_code == 1
        # 에러 출력이 있는지 확인 (stderr로 출력됨)
        assert result.output != ""

    @patch("pyhub_office_automation.excel.workbook_open.xw")
    def test_successful_open_workbook_basic(self, mock_xw, tmp_path):
        """정상적인 워크북 열기 - 기본 테스트"""
        # xlwings 모킹 설정
        mock_app = Mock()
        mock_app.visible = True
        mock_xw.App.return_value = mock_app

        mock_book = Mock()
        mock_book.name = "test_workbook.xlsx"
        mock_book.fullname = str(tmp_path / "test_workbook.xlsx")
        mock_book.saved = True

        mock_sheet = Mock()
        mock_sheet.name = "Sheet1"
        mock_sheet.index = 1
        mock_sheet.visible = True

        # 간단한 used_range 설정
        mock_used_range = Mock()
        mock_used_range.last_cell.address = "A1"
        mock_used_range.rows.count = 1
        mock_used_range.columns.count = 1
        mock_sheet.used_range = mock_used_range

        # sheets 컬렉션 모킹
        mock_sheets = Mock()
        mock_sheets.__iter__ = lambda self: iter([mock_sheet])
        mock_sheets.__len__ = lambda self: 1
        mock_sheets.active = mock_sheet
        mock_book.sheets = mock_sheets

        mock_app.books.open.return_value = mock_book

        # 임시 파일 생성
        temp_file = tmp_path / "test.xlsx"
        temp_file.touch()

        runner = CliRunner()
        result = runner.invoke(excel_app, ["workbook-open", "--file-path", str(temp_file), "--format", "json"])

        assert result.exit_code == 0

        # JSON 출력 파싱
        output_data = json.loads(result.output)

        assert output_data["success"] is True
        assert output_data["command"] == "workbook-open"
        assert "version" in output_data
        assert output_data["file_info"]["exists"] is True
