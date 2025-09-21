"""
Excel 데이터 읽기/쓰기 명령어 테스트
"""

import csv
import json
import tempfile
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest
from typer.testing import CliRunner

from pyhub_office_automation.cli.main import excel_app
from pyhub_office_automation.excel.utils import (
    create_error_response,
    create_success_response,
    load_data_from_file,
    parse_range,
    validate_range_string,
)


class TestUtilityFunctions:
    """유틸리티 함수 테스트"""

    def test_parse_range_with_sheet(self):
        """시트명 포함 범위 파싱 테스트"""
        sheet, range_part = parse_range("Sheet1!A1:C10")
        assert sheet == "Sheet1"
        assert range_part == "A1:C10"

    def test_parse_range_without_sheet(self):
        """시트명 없는 범위 파싱 테스트"""
        sheet, range_part = parse_range("A1:C10")
        assert sheet is None
        assert range_part == "A1:C10"

    def test_validate_range_string_valid(self):
        """유효한 범위 문자열 검증"""
        assert validate_range_string("A1") is True
        assert validate_range_string("A1:C10") is True
        assert validate_range_string("Sheet1!A1:C10") is True
        assert validate_range_string("AA100:ZZ999") is True

    def test_validate_range_string_invalid(self):
        """유효하지 않은 범위 문자열 검증"""
        assert validate_range_string("") is False
        assert validate_range_string("A") is False
        assert validate_range_string("1") is False
        assert validate_range_string("A1:") is False

    def test_create_error_response(self):
        """에러 응답 생성 테스트"""
        error = ValueError("Test error")
        response = create_error_response(error, "test-command")

        assert response["success"] is False
        assert response["error_type"] == "ValueError"
        assert response["error"] == "Test error"
        assert response["command"] == "test-command"
        assert "version" in response

    def test_create_success_response(self):
        """성공 응답 생성 테스트"""
        data = {"test": "data"}
        response = create_success_response(data, "test-command", "Success message")

        assert response["success"] is True
        assert response["command"] == "test-command"
        assert response["message"] == "Success message"
        assert response["data"] == data
        assert "version" in response

    def test_load_data_from_file_json(self):
        """JSON 파일 로드 테스트"""
        test_data = {"name": "John", "age": 30}

        with tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False) as f:
            json.dump(test_data, f)
            temp_path = f.name

        try:
            loaded_data = load_data_from_file(temp_path)
            assert loaded_data == test_data
        finally:
            Path(temp_path).unlink()

    def test_load_data_from_file_csv(self):
        """CSV 파일 로드 테스트"""
        test_data = [["Name", "Age"], ["John", "30"], ["Jane", "25"]]

        with tempfile.NamedTemporaryFile(mode="w", suffix=".csv", delete=False, newline="") as f:
            writer = csv.writer(f)
            writer.writerows(test_data)
            temp_path = f.name

        try:
            loaded_data = load_data_from_file(temp_path)
            assert loaded_data == test_data
        finally:
            Path(temp_path).unlink()

    def test_load_data_from_file_not_found(self):
        """존재하지 않는 파일 로드 테스트"""
        with pytest.raises(FileNotFoundError):
            load_data_from_file("non_existent_file.json")


class TestCliCommands:
    """CLI 명령어 테스트"""

    def setup_method(self):
        """테스트 설정"""
        self.runner = CliRunner()

    @patch("pyhub_office_automation.excel.range_read.get_or_open_workbook")
    def test_read_range_file_not_found(self, mock_get_or_open_workbook):
        """존재하지 않는 파일 읽기 테스트"""
        mock_get_or_open_workbook.side_effect = FileNotFoundError("File not found")

        result = self.runner.invoke(excel_app, ["range-read", "--file-path", "non_existent.xlsx", "--range", "A1:C10"])

        assert result.exit_code == 1
        response = json.loads(result.output)
        assert response["success"] is False
        assert response["error_type"] == "FileNotFoundError"

    def test_read_range_invalid_range(self):
        """잘못된 범위 형식 테스트"""
        result = self.runner.invoke(excel_app, ["range-read", "--file-path", "test.xlsx", "--range", "invalid_range"])

        assert result.exit_code == 1
        response = json.loads(result.output)
        assert response["success"] is False
        assert response["error_type"] == "ValueError"
        assert "잘못된 범위 형식" in response["error"]

    def test_write_range_no_data(self):
        """데이터 없이 쓰기 명령 테스트"""
        result = self.runner.invoke(excel_app, ["range-write", "--file-path", "test.xlsx", "--range", "A1"])

        assert result.exit_code == 1
        response = json.loads(result.output)
        assert response["success"] is False
        assert "data-file 또는 --data 중 하나를 지정해야 합니다" in response["error"]

    def test_write_range_both_data_sources(self):
        """데이터 소스 중복 지정 테스트"""
        result = self.runner.invoke(
            excel_app,
            ["range-write", "--file-path", "test.xlsx", "--range", "A1", "--data-file", "data.json", "--data", '["test"]'],
        )

        assert result.exit_code == 1
        response = json.loads(result.output)
        assert response["success"] is False
        assert "동시에 사용할 수 없습니다" in response["error"]

    def test_write_range_invalid_json(self):
        """잘못된 JSON 데이터 테스트"""
        result = self.runner.invoke(
            excel_app, ["range-write", "--file-path", "test.xlsx", "--range", "A1", "--data", "invalid_json"]
        )

        assert result.exit_code == 1
        response = json.loads(result.output)
        assert response["success"] is False
        assert "JSON 파싱 오류" in response["error"]

    def test_read_table_help(self):
        """read-table 도움말 테스트"""
        result = self.runner.invoke(excel_app, ["table-read", "--help"])
        assert result.exit_code == 0
        assert "Excel 테이블 데이터를 pandas DataFrame으로 읽습니다" in result.output

    def test_write_table_help(self):
        """write-table 도움말 테스트"""
        result = self.runner.invoke(excel_app, ["table-write", "--help"])
        assert result.exit_code == 0
        assert "pandas DataFrame을 Excel 테이블로 씁니다" in result.output

    def test_write_table_missing_data_file(self):
        """데이터 파일 누락 테스트"""
        result = self.runner.invoke(excel_app, ["table-write", "--file-path", "test.xlsx", "--data-file", "non_existent.csv"])

        assert result.exit_code == 1
        response = json.loads(result.output)
        assert response["success"] is False
        assert response["error_type"] == "FileNotFoundError"


class TestCommandIntegration:
    """명령어 통합 테스트"""

    @pytest.mark.integration
    def test_data_workflow_simulation(self):
        """데이터 워크플로우 시뮬레이션 (실제 Excel 없이)"""
        # 이 테스트는 실제 Excel이 설치된 환경에서 실행될 때만 의미가 있음
        # CI/CD에서는 skip될 수 있도록 mark 설정
        runner = CliRunner()

        # 임시 데이터 파일 생성
        test_data = [["Name", "Age", "City"], ["John", "30", "Seoul"], ["Jane", "25", "Busan"]]

        with tempfile.NamedTemporaryFile(mode="w", suffix=".csv", delete=False, newline="") as f:
            writer = csv.writer(f)
            writer.writerows(test_data)
            temp_csv = f.name

        try:
            # 1. CSV 파일이 올바르게 생성되었는지 확인
            loaded_data = load_data_from_file(temp_csv)
            assert loaded_data == test_data

            # 2. 명령어 파라미터 유효성 검증
            # (실제 Excel 없이도 파라미터 검증은 가능)

            # read-range 파라미터 검증
            assert validate_range_string("A1:C10")
            assert not validate_range_string("invalid")

            # 범위 파싱 검증
            sheet, range_part = parse_range("Data!A1:C4")
            assert sheet == "Data"
            assert range_part == "A1:C4"

        finally:
            Path(temp_csv).unlink()


if __name__ == "__main__":
    pytest.main([__file__])
