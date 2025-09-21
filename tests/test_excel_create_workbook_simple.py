"""
Excel workbook-create 명령어 간단 테스트
"""

import json
from pathlib import Path
from unittest.mock import Mock, patch

import pytest
from typer.testing import CliRunner

from pyhub_office_automation.cli.main import excel_app


class TestCreateWorkbookSimple:
    """Excel workbook-create 명령어 간단 테스트 클래스"""

    def test_help_option(self):
        """도움말 옵션 테스트"""
        runner = CliRunner()

        result = runner.invoke(excel_app, ["workbook-create", "--help"])

        assert result.exit_code == 0
        assert "새로운 Excel 워크북을 생성합니다" in result.output
        assert "--name" in result.output
        assert "--save-path" in result.output
        assert "--visible" in result.output
        assert "--format" in result.output

    def test_version_option(self):
        """버전 옵션 테스트 - Typer에서는 개별 명령어에 --version이 없음"""
        runner = CliRunner()

        # Typer에서는 개별 명령어에 --version이 없으므로 실패해야 함
        result = runner.invoke(excel_app, ["workbook-create", "--version"])

        assert result.exit_code != 0  # 실패해야 함
        assert "No such option" in result.output or "Unrecognized arguments" in result.output

    @patch("pyhub_office_automation.excel.workbook_create.xw")
    def test_successful_create_workbook_basic(self, mock_xw):
        """정상적인 워크북 생성 - 기본 테스트"""
        # xlwings 모킹 설정
        mock_app = Mock()
        mock_app.visible = True
        mock_xw.App.return_value = mock_app

        mock_book = Mock()
        mock_book.name = "Book1"
        mock_book.fullname = "Book1"
        mock_book.saved = False

        mock_sheet = Mock()
        mock_sheet.name = "Sheet1"
        mock_sheet.index = 1
        mock_sheet.visible = True

        # sheets 컬렉션 모킹
        mock_sheets = Mock()
        mock_sheets.__iter__ = lambda self: iter([mock_sheet])
        mock_sheets.__len__ = lambda self: 1
        mock_sheets.active = mock_sheet
        mock_book.sheets = mock_sheets

        mock_app.books.add.return_value = mock_book

        runner = CliRunner()
        result = runner.invoke(excel_app, ["workbook-create", "--name", "TestWorkbook", "--format", "json"])

        if result.exit_code != 0:
            print(f"Error output: {result.output}")
            if result.exception:
                print(f"Exception: {result.exception}")

        assert result.exit_code == 0

        # JSON 출력 파싱
        output_data = json.loads(result.output)

        assert output_data["success"] is True
        assert output_data["command"] == "workbook-create"
        assert "version" in output_data
        assert output_data["workbook_info"]["name"] == "Book1"

    @patch("pyhub_office_automation.excel.workbook_create.xw")
    def test_excel_application_error(self, mock_xw):
        """Excel 애플리케이션 시작 실패 테스트"""
        # Excel 애플리케이션 시작 실패 설정
        mock_xw.App.side_effect = Exception("Excel을 시작할 수 없습니다")

        runner = CliRunner()
        result = runner.invoke(excel_app, ["workbook-create", "--name", "TestWorkbook", "--format", "json"])

        assert result.exit_code == 1

        # JSON 에러 출력 파싱
        output_data = json.loads(result.output)

        assert output_data["success"] is False
        assert output_data["error_type"] == "RuntimeError"
        assert "Excel 애플리케이션을 시작할 수 없습니다" in output_data["error"]
