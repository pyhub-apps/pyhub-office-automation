"""
Excel 시트 관리 명령어들 테스트
add-sheet, rename-sheet, activate-sheet, delete-sheet 테스트
"""

import json
from pathlib import Path
from unittest.mock import MagicMock, Mock, patch

import pytest
from typer.testing import CliRunner

from pyhub_office_automation.cli.main import excel_app


def create_mock_sheets_collection(sheet_list, active_sheet=None):
    """Mock sheets collection helper"""
    mock_sheets = Mock()
    mock_sheets.__iter__ = Mock(return_value=iter(sheet_list))
    mock_sheets.__len__ = Mock(return_value=len(sheet_list))
    mock_sheets.__getitem__ = Mock(
        side_effect=lambda x: sheet_list[x] if isinstance(x, int) else next(s for s in sheet_list if s.name == x)
    )
    mock_sheets.active = active_sheet or (sheet_list[0] if sheet_list else None)
    mock_sheets.add = Mock()
    return mock_sheets


class TestAddSheet:
    """Excel add-sheet 명령어 테스트 클래스"""

    def test_help_option(self):
        """도움말 옵션 테스트"""
        runner = CliRunner()
        result = runner.invoke(excel_app, ["sheet-add", "--help"])

        assert result.exit_code == 0
        assert "Excel 워크북에 새 워크시트를 추가합니다" in result.output
        assert "--workbook" in result.output
        assert "--name" in result.output
        assert "--before" in result.output
        assert "--after" in result.output

    @patch("pyhub_office_automation.excel.sheet_add.get_or_open_workbook")
    def test_successful_add_sheet_basic(self, mock_get_or_open_workbook):
        """정상적인 시트 추가 - 기본 테스트"""
        # 모킹 설정
        mock_book = Mock()
        mock_book.name = "TestBook.xlsx"
        mock_book.fullname = "/path/to/TestBook.xlsx"

        # 기존 시트들
        mock_sheet1 = Mock()
        mock_sheet1.name = "Sheet1"
        mock_sheet1.index = 1

        mock_sheet2 = Mock()
        mock_sheet2.name = "NewSheet"
        mock_sheet2.index = 2
        mock_sheet2.visible = True

        # sheets 컬렉션 모킹
        mock_sheets = create_mock_sheets_collection([mock_sheet1], mock_sheet1)
        mock_sheets.add.return_value = mock_sheet2
        # 새 시트 추가 후 상태 업데이트
        mock_sheets.__iter__ = Mock(return_value=iter([mock_sheet1, mock_sheet2]))
        mock_sheets.__len__ = Mock(return_value=2)

        mock_book.sheets = mock_sheets
        mock_book.save = Mock()

        mock_get_or_open_workbook.return_value = mock_book

        runner = CliRunner()
        result = runner.invoke(excel_app, ["sheet-add", "--file-path", "test.xlsx", "--name", "NewSheet"])

        assert result.exit_code == 0
        response = json.loads(result.output)
        assert response["success"] is True
        assert response["command"] == "add-sheet"
        assert response["data"]["new_sheet"]["name"] == "NewSheet"

    def test_add_sheet_duplicate_name_error(self):
        """중복 이름 에러 테스트"""
        with patch("pyhub_office_automation.excel.sheet_add.get_or_open_workbook") as mock_get_or_open_workbook:
            mock_book = Mock()
            mock_sheet1 = Mock()
            mock_sheet1.name = "ExistingSheet"

            mock_sheets = create_mock_sheets_collection([mock_sheet1], mock_sheet1)
            mock_book.sheets = mock_sheets
            mock_get_or_open_workbook.return_value = mock_book

            runner = CliRunner()
            result = runner.invoke(excel_app, ["sheet-add", "--file-path", "test.xlsx", "--name", "ExistingSheet"])

            assert result.exit_code == 1
            response = json.loads(result.output)
            assert response["success"] is False
            assert "이미 존재합니다" in response["error"]

    def test_add_sheet_invalid_options(self):
        """잘못된 옵션 조합 테스트"""
        runner = CliRunner()
        result = runner.invoke(
            excel_app, ["sheet-add", "--file-path", "test.xlsx", "--before", "Sheet1", "--after", "Sheet2", "--index", "1"]
        )

        assert result.exit_code == 1
        response = json.loads(result.output)
        assert response["success"] is False
        assert "중 하나만 지정" in response["error"]


class TestRenameSheet:
    """Excel rename-sheet 명령어 테스트 클래스"""

    def test_help_option(self):
        """도움말 옵션 테스트"""
        runner = CliRunner()
        result = runner.invoke(excel_app, ["sheet-rename", "--help"])

        assert result.exit_code == 0
        assert "Excel 워크북의 시트 이름을 변경합니다" in result.output
        assert "--old-name" in result.output
        assert "--new-name" in result.output

    @patch("pyhub_office_automation.excel.sheet_rename.get_or_open_workbook")
    def test_successful_rename_sheet(self, mock_get_or_open_workbook):
        """정상적인 시트 이름 변경 테스트"""
        mock_book = Mock()
        mock_sheet = Mock()
        mock_sheet.name = "OldName"
        mock_sheet.index = 1
        mock_sheet.visible = True

        mock_sheets = create_mock_sheets_collection([mock_sheet], mock_sheet)
        mock_book.sheets = mock_sheets
        mock_book.save = Mock()

        # 시트 이름 변경 시뮬레이션
        def set_name(value):
            mock_sheet.name = value

        type(mock_sheet).name = property(lambda self: mock_sheet.name, set_name)

        mock_get_or_open_workbook.return_value = mock_book

        runner = CliRunner()
        result = runner.invoke(
            excel_app, ["sheet-rename", "--file-path", "test.xlsx", "--old-name", "OldName", "--new-name", "NewName"]
        )

        assert result.exit_code == 0
        response = json.loads(result.output)
        assert response["success"] is True

    def test_rename_sheet_invalid_chars(self):
        """잘못된 문자 포함 이름 테스트"""
        runner = CliRunner()
        result = runner.invoke(
            excel_app, ["sheet-rename", "--file-path", "test.xlsx", "--old-name", "Sheet1", "--new-name", "Invalid/Name"]
        )

        assert result.exit_code == 1
        response = json.loads(result.output)
        assert response["success"] is False
        assert "사용할 수 없는 문자" in response["error"]

    def test_rename_sheet_too_long_name(self):
        """너무 긴 이름 테스트"""
        long_name = "A" * 32  # 31자 초과
        runner = CliRunner()
        result = runner.invoke(
            excel_app, ["sheet-rename", "--file-path", "test.xlsx", "--old-name", "Sheet1", "--new-name", long_name]
        )

        assert result.exit_code == 1
        response = json.loads(result.output)
        assert response["success"] is False
        assert "31자를 초과" in response["error"]


class TestActivateSheet:
    """Excel activate-sheet 명령어 테스트 클래스"""

    def test_help_option(self):
        """도움말 옵션 테스트"""
        runner = CliRunner()
        result = runner.invoke(excel_app, ["sheet-activate", "--help"])

        assert result.exit_code == 0
        assert "Excel 워크북의 특정 시트를 활성화합니다" in result.output
        assert "--name" in result.output
        assert "--index" in result.output

    @patch("pyhub_office_automation.excel.sheet_activate.get_or_open_workbook")
    def test_successful_activate_sheet(self, mock_get_or_open_workbook):
        """정상적인 시트 활성화 테스트"""
        mock_book = Mock()
        mock_sheet1 = Mock()
        mock_sheet1.name = "Sheet1"
        mock_sheet1.index = 1
        mock_sheet1.visible = True
        mock_sheet1.activate = Mock()

        mock_sheet2 = Mock()
        mock_sheet2.name = "Sheet2"

        mock_sheets = create_mock_sheets_collection([mock_sheet1, mock_sheet2], mock_sheet1)
        mock_book.sheets = mock_sheets

        mock_get_or_open_workbook.return_value = mock_book

        runner = CliRunner()
        result = runner.invoke(excel_app, ["sheet-activate", "--file-path", "test.xlsx", "--name", "Sheet1"])

        assert result.exit_code == 0
        response = json.loads(result.output)
        assert response["success"] is True
        assert response["data"]["activated_sheet"]["name"] == "Sheet1"
        mock_sheet1.activate.assert_called_once()

    def test_activate_sheet_not_found(self):
        """존재하지 않는 시트 활성화 테스트"""
        with patch("pyhub_office_automation.excel.sheet_activate.get_or_open_workbook") as mock_get_or_open_workbook:
            mock_book = Mock()
            mock_sheet = Mock()
            mock_sheet.name = "Sheet1"

            mock_sheets = create_mock_sheets_collection([mock_sheet], mock_sheet)
            mock_book.sheets = mock_sheets
            mock_get_or_open_workbook.return_value = mock_book

            runner = CliRunner()
            result = runner.invoke(excel_app, ["sheet-activate", "--file-path", "test.xlsx", "--name", "NonExistentSheet"])

            assert result.exit_code == 1
            response = json.loads(result.output)
            assert response["success"] is False
            assert "찾을 수 없습니다" in response["error"]


class TestDeleteSheet:
    """Excel delete-sheet 명령어 테스트 클래스"""

    def test_help_option(self):
        """도움말 옵션 테스트"""
        runner = CliRunner()
        result = runner.invoke(excel_app, ["sheet-delete", "--help"])

        assert result.exit_code == 0
        assert "Excel 워크북에서 시트를 삭제합니다" in result.output
        assert "--name" in result.output
        assert "--force" in result.output

    @patch("pyhub_office_automation.excel.sheet_delete.get_or_open_workbook")
    def test_successful_delete_sheet(self, mock_get_or_open_workbook):
        """정상적인 시트 삭제 테스트"""
        mock_book = Mock()

        # 2개 시트 생성 (최소 1개는 남아야 함)
        mock_sheet1 = Mock()
        mock_sheet1.name = "Sheet1"
        mock_sheet1.index = 1
        mock_sheet1.visible = True
        mock_sheet1.delete = Mock()

        mock_sheet2 = Mock()
        mock_sheet2.name = "Sheet2"
        mock_sheet2.activate = Mock()

        mock_sheets = create_mock_sheets_collection([mock_sheet1, mock_sheet2], mock_sheet1)
        mock_book.sheets = mock_sheets
        mock_book.save = Mock()

        # 삭제 후 상태 시뮬레이션
        def delete_side_effect():
            mock_sheets.__iter__ = Mock(return_value=iter([mock_sheet2]))
            mock_sheets.__len__ = Mock(return_value=1)
            mock_sheets.active = mock_sheet2

        mock_sheet1.delete.side_effect = delete_side_effect

        mock_get_or_open_workbook.return_value = mock_book

        runner = CliRunner()
        result = runner.invoke(
            excel_app, ["sheet-delete", "--file-path", "test.xlsx", "--name", "Sheet1", "--force"]  # 확인 없이 삭제
        )

        assert result.exit_code == 0
        response = json.loads(result.output)
        assert response["success"] is True
        assert response["data"]["deleted_sheet"]["name"] == "Sheet1"
        mock_sheet1.delete.assert_called_once()

    def test_delete_last_sheet_error(self):
        """마지막 시트 삭제 시도 에러 테스트"""
        with patch("pyhub_office_automation.excel.sheet_delete.get_or_open_workbook") as mock_get_or_open_workbook:
            mock_book = Mock()
            mock_sheet = Mock()
            mock_sheet.name = "Sheet1"

            mock_sheets = create_mock_sheets_collection([mock_sheet], mock_sheet)
            mock_book.sheets = mock_sheets
            mock_get_or_open_workbook.return_value = mock_book

            runner = CliRunner()
            result = runner.invoke(excel_app, ["sheet-delete", "--file-path", "test.xlsx", "--name", "Sheet1", "--force"])

            assert result.exit_code == 1
            response = json.loads(result.output)
            assert response["success"] is False
            assert "최소 1개의 시트가 필요" in response["error"]

    def test_delete_sheet_not_found(self):
        """존재하지 않는 시트 삭제 테스트"""
        with patch("pyhub_office_automation.excel.sheet_delete.get_or_open_workbook") as mock_get_or_open_workbook:
            mock_book = Mock()
            mock_sheet = Mock()
            mock_sheet.name = "Sheet1"

            # 2개 시트로 설정 (삭제 가능한 상태)
            mock_sheet2 = Mock()
            mock_sheet2.name = "Sheet2"
            mock_sheets = create_mock_sheets_collection([mock_sheet, mock_sheet2], mock_sheet)
            mock_book.sheets = mock_sheets
            mock_get_or_open_workbook.return_value = mock_book

            runner = CliRunner()
            result = runner.invoke(
                excel_app, ["sheet-delete", "--file-path", "test.xlsx", "--name", "NonExistentSheet", "--force"]
            )

            assert result.exit_code == 1
            response = json.loads(result.output)
            assert response["success"] is False
            assert "찾을 수 없습니다" in response["error"]


class TestSheetManagementEdgeCases:
    """시트 관리 명령어들의 엣지 케이스 테스트"""

    def test_missing_workbook_option(self):
        """워크북 옵션 누락 테스트"""
        runner = CliRunner()

        # add-sheet 테스트
        result = runner.invoke(excel_app, ["sheet-add", "--name", "TestSheet"])
        assert result.exit_code != 0
        assert "Missing option" in result.output

    def test_missing_required_options(self):
        """필수 옵션 누락 테스트"""
        runner = CliRunner()

        # rename-sheet에서 new-name 누락
        result = runner.invoke(excel_app, ["sheet-rename", "--file-path", "test.xlsx", "--old-name", "Sheet1"])
        assert result.exit_code != 0

    def test_conflicting_options(self):
        """상충하는 옵션 테스트"""
        runner = CliRunner()

        # activate-sheet에서 name과 index 동시 사용
        result = runner.invoke(excel_app, ["sheet-activate", "--file-path", "test.xlsx", "--name", "Sheet1", "--index", "0"])
        assert result.exit_code == 1

    @patch("pyhub_office_automation.excel.sheet_add.get_or_open_workbook")
    def test_text_output_format(self, mock_get_or_open_workbook):
        """텍스트 출력 형식 테스트"""
        mock_book = Mock()
        mock_sheet = Mock()
        mock_sheet.name = "NewSheet"
        mock_sheet.index = 1
        mock_sheet.visible = True

        mock_sheets = create_mock_sheets_collection([], None)
        mock_sheets.add.return_value = mock_sheet
        # 새 시트 추가 후 상태
        mock_sheets.__iter__ = Mock(return_value=iter([mock_sheet]))
        mock_sheets.__len__ = Mock(return_value=1)
        mock_sheets.active = mock_sheet

        mock_book.sheets = mock_sheets
        mock_book.save = Mock()
        mock_get_or_open_workbook.return_value = mock_book

        runner = CliRunner()
        result = runner.invoke(excel_app, ["sheet-add", "--file-path", "test.xlsx", "--name", "NewSheet", "--format", "text"])

        assert result.exit_code == 0
        assert "✅ 시트 추가 성공" in result.output
        # JSON이 아닌 텍스트 형식인지 확인
        assert not result.output.strip().startswith("{")
