"""
Excel workbook-open 명령어 테스트
"""

import json
from pathlib import Path
from unittest.mock import Mock, patch

import pytest
from typer.testing import CliRunner

from pyhub_office_automation.cli.main import excel_app


class TestWorkbookOpen:
    """Excel workbook-open 명령어 테스트 클래스"""

    @patch("pyhub_office_automation.excel.workbook_open.xw")
    def test_successful_open_workbook_json_output(self, mock_xw, temp_excel_file):
        """정상적인 워크북 열기 - JSON 출력 테스트"""
        # xlwings 모킹 설정
        mock_app = Mock()
        mock_app.visible = True
        mock_xw.App.return_value = mock_app

        mock_book = Mock()
        mock_book.name = "test_workbook.xlsx"
        mock_book.fullname = "/path/to/test_workbook.xlsx"
        mock_book.saved = True

        mock_sheet = Mock()
        mock_sheet.name = "Sheet1"
        mock_sheet.index = 1
        mock_sheet.visible = True

        mock_used_range = Mock()
        mock_used_range.last_cell.address = "C5"
        mock_used_range.rows.count = 5
        mock_used_range.columns.count = 3
        mock_sheet.used_range = mock_used_range

        mock_sheets = Mock()
        mock_sheets.__iter__ = lambda self: iter([mock_sheet])
        mock_sheets.__len__ = lambda self: 1
        mock_sheets.active = mock_sheet
        mock_book.sheets = mock_sheets

        mock_app.books.open.return_value = mock_book

        runner = CliRunner()

        # 실제 파일이 존재하도록 생성
        temp_excel_file.touch()

        result = runner.invoke(excel_app, ["workbook-open", "--file-path", str(temp_excel_file), "--format", "json"])

        assert result.exit_code == 0

        # JSON 출력 파싱
        output_data = json.loads(result.output)

        assert output_data["success"] is True
        assert output_data["command"] == "workbook-open"
        assert "version" in output_data
        assert output_data["file_info"]["exists"] is True
        assert output_data["file_info"]["name"] == temp_excel_file.name
        assert output_data["workbook_info"]["name"] == "test_workbook.xlsx"
        assert len(output_data["sheets"]) == 1
        assert output_data["sheets"][0]["name"] == "Sheet1"

    def test_successful_open_workbook_text_output(self, temp_excel_file, mock_xlwings):
        """정상적인 워크북 열기 - 텍스트 출력 테스트"""
        runner = CliRunner()

        # 실제 파일이 존재하도록 생성
        temp_excel_file.touch()

        result = runner.invoke(excel_app, ["workbook-open", "--file-path", str(temp_excel_file), "--format", "text"])

        assert result.exit_code == 0
        assert "✅ 워크북 열기 성공" in result.output
        assert "📊 시트 수: 1" in result.output
        assert "🎯 활성 시트: Sheet1" in result.output

    def test_file_not_found_error(self, non_existent_file):
        """파일이 존재하지 않는 경우 테스트"""
        runner = CliRunner()

        result = runner.invoke(excel_app, ["workbook-open", "--file-path", non_existent_file, "--format", "json"])

        assert result.exit_code == 1

        # JSON 에러 출력 파싱
        output_data = json.loads(result.output)

        assert output_data["success"] is False
        assert output_data["error_type"] == "FileNotFoundError"
        assert "FileNotFoundError" in output_data["error"] or "파일을 찾을 수 없습니다" in output_data["error"]

    def test_invalid_file_extension(self, temp_invalid_file):
        """잘못된 파일 확장자 테스트"""
        runner = CliRunner()

        result = runner.invoke(excel_app, ["workbook-open", "--file-path", str(temp_invalid_file), "--format", "json"])

        assert result.exit_code == 1

        # JSON 에러 출력 파싱
        output_data = json.loads(result.output)

        assert output_data["success"] is False
        assert output_data["error_type"] == "ValueError"
        assert "지원되지 않는 파일 형식" in output_data["error"]

    def test_excel_application_error(self, temp_excel_file, mock_xlwings_error):
        """Excel 애플리케이션 시작 실패 테스트"""
        runner = CliRunner()

        # 실제 파일이 존재하도록 생성
        temp_excel_file.touch()

        result = runner.invoke(excel_app, ["workbook-open", "--file-path", str(temp_excel_file), "--format", "json"])

        assert result.exit_code == 1

        # JSON 에러 출력 파싱
        output_data = json.loads(result.output)

        assert output_data["success"] is False
        assert output_data["error_type"] == "RuntimeError"
        assert "Excel 애플리케이션을 시작할 수 없습니다" in output_data["error"]

    def test_workbook_open_error(self, temp_excel_file, mock_xlwings):
        """워크북 열기 실패 테스트"""
        runner = CliRunner()

        # 실제 파일이 존재하도록 생성
        temp_excel_file.touch()

        # 워크북 열기 실패 설정
        mock_xlwings["app"].books.open.side_effect = Exception("워크북을 열 수 없습니다")

        result = runner.invoke(excel_app, ["workbook-open", "--file-path", str(temp_excel_file), "--format", "json"])

        assert result.exit_code == 1

        # JSON 에러 출력 파싱
        output_data = json.loads(result.output)

        assert output_data["success"] is False
        assert output_data["error_type"] == "RuntimeError"
        assert "워크북을 열 수 없습니다" in output_data["error"]

    def test_visible_option(self, temp_excel_file, mock_xlwings):
        """visible 옵션 테스트"""
        runner = CliRunner()

        # 실제 파일이 존재하도록 생성
        temp_excel_file.touch()

        result = runner.invoke(
            excel_app, ["workbook-open", "--file-path", str(temp_excel_file), "--visible", "False", "--format", "json"]
        )

        assert result.exit_code == 0

        # xlwings App이 visible=False로 호출되었는지 확인
        mock_xlwings["app_class"].assert_called_with(visible=False)

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

    def test_empty_sheet_handling(self, temp_excel_file, mock_xlwings):
        """빈 시트 처리 테스트"""
        runner = CliRunner()

        # 실제 파일이 존재하도록 생성
        temp_excel_file.touch()

        # 빈 시트 설정 (used_range가 None인 경우)
        mock_xlwings["sheet"].used_range = None

        result = runner.invoke(excel_app, ["workbook-open", "--file-path", str(temp_excel_file), "--format", "json"])

        assert result.exit_code == 0

        output_data = json.loads(result.output)
        assert output_data["success"] is True
        assert output_data["sheets"][0]["used_range"]["last_cell"] == "A1"
        assert output_data["sheets"][0]["used_range"]["row_count"] == 0
        assert output_data["sheets"][0]["used_range"]["column_count"] == 0

    def test_sheet_info_collection_error(self, temp_excel_file, mock_xlwings):
        """시트 정보 수집 실패 처리 테스트"""
        runner = CliRunner()

        # 실제 파일이 존재하도록 생성
        temp_excel_file.touch()

        # 시트 정보 수집 시 에러 발생 설정
        mock_xlwings["sheet"].used_range = property(lambda self: (_ for _ in ()).throw(Exception("시트 접근 오류")))

        result = runner.invoke(excel_app, ["workbook-open", "--file-path", str(temp_excel_file), "--format", "json"])

        assert result.exit_code == 0

        output_data = json.loads(result.output)
        assert output_data["success"] is True
        # 에러가 발생한 시트는 error 필드를 포함해야 함
        assert "error" in output_data["sheets"][0]
