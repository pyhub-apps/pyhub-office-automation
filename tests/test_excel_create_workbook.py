"""
Excel create-workbook 명령어 테스트
"""

import json
import pytest
from pathlib import Path
from click.testing import CliRunner
from unittest.mock import patch, Mock

from pyhub_office_automation.excel.create_workbook import create_workbook


class TestCreateWorkbook:
    """Excel create-workbook 명령어 테스트 클래스"""

    def test_help_option(self):
        """도움말 옵션 테스트"""
        runner = CliRunner()

        result = runner.invoke(create_workbook, ['--help'])

        assert result.exit_code == 0
        assert '새로운 Excel 워크북을 생성합니다' in result.output
        assert '--name' in result.output
        assert '--save-path' in result.output
        assert '--visible' in result.output
        assert '--format' in result.output

    def test_version_option(self):
        """버전 옵션 테스트"""
        runner = CliRunner()

        result = runner.invoke(create_workbook, ['--version'])

        assert result.exit_code == 0
        # 버전 정보가 출력되는지 확인
        assert result.output.strip() != ""

    @patch('pyhub_office_automation.excel.create_workbook.xw')
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
        result = runner.invoke(create_workbook, [
            '--name', 'TestWorkbook',
            '--format', 'json'
        ])

        assert result.exit_code == 0

        # JSON 출력 파싱
        output_data = json.loads(result.output)

        assert output_data['success'] is True
        assert output_data['command'] == 'create-workbook'
        assert 'version' in output_data
        assert output_data['workbook_info']['name'] == 'Book1'
        assert output_data['workbook_info']['saved'] is False
        assert len(output_data['sheets']) == 1
        assert output_data['sheets'][0]['name'] == 'Sheet1'

    @patch('pyhub_office_automation.excel.create_workbook.xw')
    def test_successful_create_workbook_text_output(self, mock_xw):
        """정상적인 워크북 생성 - 텍스트 출력 테스트"""
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
        result = runner.invoke(create_workbook, [
            '--name', 'TestWorkbook',
            '--format', 'text'
        ])

        assert result.exit_code == 0
        assert "✅ 새 워크북 생성 성공" in result.output
        assert "📊 시트 수: 1" in result.output
        assert "🎯 활성 시트: Sheet1" in result.output
        assert "📝 저장되지 않음" in result.output

    @patch('pyhub_office_automation.excel.create_workbook.xw')
    def test_create_workbook_with_save_path(self, mock_xw, tmp_path):
        """저장 경로 지정한 워크북 생성 테스트"""
        # xlwings 모킹 설정
        mock_app = Mock()
        mock_app.visible = True
        mock_xw.App.return_value = mock_app

        mock_book = Mock()
        mock_book.name = "TestWorkbook.xlsx"
        mock_book.fullname = str(tmp_path / "TestWorkbook.xlsx")
        mock_book.saved = True

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

        save_path = tmp_path / "TestWorkbook.xlsx"

        runner = CliRunner()
        result = runner.invoke(create_workbook, [
            '--name', 'TestWorkbook',
            '--save-path', str(save_path),
            '--format', 'json'
        ])

        assert result.exit_code == 0

        # JSON 출력 파싱
        output_data = json.loads(result.output)

        assert output_data['success'] is True
        assert output_data['workbook_info']['saved'] is True
        assert output_data['workbook_info']['saved_path'] == str(save_path)

        # save 메서드가 호출되었는지 확인
        mock_book.save.assert_called_once()

    @patch('pyhub_office_automation.excel.create_workbook.xw')
    def test_excel_application_error(self, mock_xw):
        """Excel 애플리케이션 시작 실패 테스트"""
        # Excel 애플리케이션 시작 실패 설정
        mock_xw.App.side_effect = Exception("Excel을 시작할 수 없습니다")

        runner = CliRunner()
        result = runner.invoke(create_workbook, [
            '--name', 'TestWorkbook',
            '--format', 'json'
        ])

        assert result.exit_code == 1

        # JSON 에러 출력 파싱
        output_data = json.loads(result.output)

        assert output_data['success'] is False
        assert output_data['error_type'] == 'RuntimeError'
        assert 'Excel 애플리케이션을 시작할 수 없습니다' in output_data['error']

    @patch('pyhub_office_automation.excel.create_workbook.xw')
    def test_workbook_creation_error(self, mock_xw):
        """워크북 생성 실패 테스트"""
        # xlwings 모킹 설정
        mock_app = Mock()
        mock_xw.App.return_value = mock_app

        # 워크북 생성 실패 설정
        mock_app.books.add.side_effect = Exception("워크북을 생성할 수 없습니다")

        runner = CliRunner()
        result = runner.invoke(create_workbook, [
            '--name', 'TestWorkbook',
            '--format', 'json'
        ])

        assert result.exit_code == 1

        # JSON 에러 출력 파싱
        output_data = json.loads(result.output)

        assert output_data['success'] is False
        assert output_data['error_type'] == 'RuntimeError'
        assert '새 워크북을 생성할 수 없습니다' in output_data['error']

    @patch('pyhub_office_automation.excel.create_workbook.xw')
    def test_visible_option(self, mock_xw):
        """visible 옵션 테스트"""
        # xlwings 모킹 설정
        mock_app = Mock()
        mock_xw.App.return_value = mock_app

        mock_book = Mock()
        mock_book.name = "Book1"
        mock_book.fullname = "Book1"
        mock_book.saved = False

        mock_sheet = Mock()
        mock_sheet.name = "Sheet1"
        mock_sheet.index = 1
        mock_sheet.visible = True

        mock_sheets = Mock()
        mock_sheets.__iter__ = lambda self: iter([mock_sheet])
        mock_sheets.__len__ = lambda self: 1
        mock_sheets.active = mock_sheet
        mock_book.sheets = mock_sheets

        mock_app.books.add.return_value = mock_book

        runner = CliRunner()
        result = runner.invoke(create_workbook, [
            '--name', 'TestWorkbook',
            '--visible', 'False',
            '--format', 'json'
        ])

        assert result.exit_code == 0

        # xlwings App이 visible=False로 호출되었는지 확인
        mock_xw.App.assert_called_with(visible=False)