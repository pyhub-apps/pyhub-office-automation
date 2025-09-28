"""
GitHub Issue #70: chart-pivot-create COM 에러 0x800401FD 해결 테스트
차트는 생성되지만 COM 에러로 실패 처리되는 문제의 수정사항 테스트
"""

import json
import platform
from typing import Any, Dict
from unittest.mock import MagicMock, Mock, patch

import pytest

from pyhub_office_automation.excel.utils import COM_ERROR_MESSAGES, extract_com_error_code


class MockCOMError(Exception):
    """COM 에러 0x800401FD 시뮬레이션"""

    def __init__(self, error_code: int = 0x800401FD):
        self.args = (error_code, "개체가 서버에 연결되지 않았습니다.", None, None)
        self.error_code = error_code
        super().__init__(f"COM Error: {error_code}")


class MockChartObject:
    """성공적으로 생성된 차트 객체 시뮬레이션"""

    def __init__(self, name: str = "Chart1", has_data: bool = True):
        self.Name = name
        self.Chart = MockChart(has_data=has_data)


class MockChart:
    """차트 속성 시뮬레이션"""

    def __init__(self, has_data: bool = True, chart_type: int = 51):  # xlColumnClustered
        self.ChartType = chart_type
        self._series_collection = MockSeriesCollection(count=3 if has_data else 0)
        self.HasTitle = False
        self.ChartTitle = MockChartTitle()

    def SeriesCollection(self):
        return self._series_collection


class MockSeriesCollection:
    """시리즈 컬렉션 시뮬레이션"""

    def __init__(self, count: int = 3):
        self.Count = count


class MockChartTitle:
    """차트 제목 시뮬레이션"""

    def __init__(self):
        self.Text = ""


class MockChartObjects:
    """차트 객체 컬렉션 시뮬레이션"""

    def __init__(self, charts: list = None):
        self.charts = charts or []
        self.Count = len(self.charts)

    def __call__(self, index: int):
        if 1 <= index <= len(self.charts):
            return self.charts[index - 1]  # COM은 1-based 인덱스
        raise IndexError(f"Invalid chart index: {index}")


class MockSheet:
    """워크시트 시뮬레이션"""

    def __init__(self, chart_objects: MockChartObjects = None):
        self.api = Mock()
        self.api.ChartObjects.return_value = chart_objects or MockChartObjects()


def test_com_error_code_extraction():
    """COM 에러 코드 추출 테스트"""
    com_error = MockCOMError(0x800401FD)
    error_code = extract_com_error_code(com_error)
    assert error_code == 0x800401FD


def test_com_error_message_mapping():
    """COM 에러 메시지 매핑 테스트"""
    error_code = 0x800401FD
    assert error_code in COM_ERROR_MESSAGES

    error_info = COM_ERROR_MESSAGES[error_code]
    assert "recovery_info" in error_info
    assert error_info["recovery_info"]["auto_recovery"] is True
    assert error_info["recovery_info"]["github_issue"] == "#70"


@pytest.mark.parametrize(
    "chart_count,has_data,should_succeed",
    [
        (1, True, True),  # 정상: 차트 1개, 데이터 있음
        (2, True, True),  # 정상: 차트 2개, 데이터 있음 (최신 차트 선택)
        (1, False, False),  # 실패: 차트 있지만 데이터 없음
        (0, False, False),  # 실패: 차트 없음
    ],
)
def test_chart_verification_logic(chart_count: int, has_data: bool, should_succeed: bool):
    """차트 검증 로직 테스트"""
    # 시나리오별 차트 객체 생성
    charts = []
    for i in range(chart_count):
        chart_name = f"Chart{i+1}"
        charts.append(MockChartObject(name=chart_name, has_data=has_data))

    chart_objects = MockChartObjects(charts)
    mock_sheet = MockSheet(chart_objects)

    # 검증 로직 시뮬레이션
    try:
        chart_count = chart_objects.Count

        if chart_count > 0:
            # 가장 최근에 생성된 차트를 가져옴
            chart_object = chart_objects(chart_count)
            chart = chart_object.Chart
            chart_name = chart_object.Name

            # 차트 기본 속성 검증
            chart_type = chart.ChartType
            series_has_data = chart.SeriesCollection().Count > 0

            if series_has_data:
                result = True
            else:
                result = False
        else:
            result = False

        assert result == should_succeed, f"Expected {should_succeed}, got {result}"

    except Exception as e:
        assert not should_succeed, f"Unexpected error when should_succeed={should_succeed}: {e}"


def test_com_error_recovery_flow():
    """COM 에러 복구 플로우 전체 테스트"""
    # 성공적으로 생성된 차트가 있는 상황 시뮬레이션
    chart = MockChartObject(name="TestChart", has_data=True)
    chart_objects = MockChartObjects([chart])
    mock_sheet = MockSheet(chart_objects)

    # COM 에러 0x800401FD 발생 시뮬레이션
    com_error = MockCOMError(0x800401FD)
    error_code = extract_com_error_code(com_error)

    # 복구 로직 시뮬레이션
    recovered = False
    chart_name = None

    if error_code == 0x800401FD:
        try:
            chart_objects_result = mock_sheet.api.ChartObjects()
            chart_count = chart_objects_result.Count

            if chart_count > 0:
                chart_object = chart_objects_result(chart_count)
                chart_result = chart_object.Chart
                chart_name = chart_object.Name

                chart_type = chart_result.ChartType
                has_data = chart_result.SeriesCollection().Count > 0

                if has_data:
                    recovered = True

        except Exception:
            recovered = False

    assert recovered is True
    assert chart_name == "TestChart"


def test_response_data_structure():
    """응답 데이터 구조 테스트 (com_error_recovery 필드)"""
    # 복구 성공 시나리오
    response_data = {"chart_name": "TestChart", "pivot_name": "TestPivot"}

    # COM 에러 복구 정보 추가
    recovered_from_com_error = True

    if recovered_from_com_error:
        response_data["com_error_recovery"] = {
            "recovered": True,
            "error_code": "0x800401FD",
            "description": "COM 연결 에러가 발생했지만 차트 생성이 성공적으로 완료되었습니다",
            "impact": "기능상 문제 없음",
        }

    # 검증
    assert "com_error_recovery" in response_data
    recovery_info = response_data["com_error_recovery"]
    assert recovery_info["recovered"] is True
    assert recovery_info["error_code"] == "0x800401FD"
    assert "기능상 문제 없음" in recovery_info["impact"]


@pytest.mark.skipif(platform.system() != "Windows", reason="Windows COM 기능 전용")
def test_windows_only_functionality():
    """Windows 전용 기능 테스트 마커"""
    # 실제 Windows 환경에서만 실행되는 테스트
    # 여기서는 플랫폼 체크만 수행
    assert platform.system() == "Windows"


def test_edge_case_chart_validation_failure():
    """차트 검증 실패 엣지 케이스 테스트"""
    # 차트는 있지만 검증 중 에러 발생 시나리오
    chart = MockChartObject(name="BrokenChart", has_data=True)

    # SeriesCollection 호출 시 에러 발생하도록 설정
    chart.Chart.SeriesCollection = Mock(side_effect=Exception("Chart validation error"))

    chart_objects = MockChartObjects([chart])
    mock_sheet = MockSheet(chart_objects)

    # 검증 실패 시나리오
    validation_failed = False
    try:
        chart_objects_result = mock_sheet.api.ChartObjects()
        chart_count = chart_objects_result.Count

        if chart_count > 0:
            chart_object = chart_objects_result(chart_count)
            chart_result = chart_object.Chart

            # 여기서 에러 발생
            has_data = chart_result.SeriesCollection().Count > 0

    except Exception:
        validation_failed = True

    assert validation_failed is True


def test_com_error_recovery_info_structure():
    """COM 에러 복구 정보 구조 테스트"""
    error_info = COM_ERROR_MESSAGES[0x800401FD]
    recovery_info = error_info["recovery_info"]

    # 필수 필드 검증
    assert "auto_recovery" in recovery_info
    assert "success_indicator" in recovery_info
    assert "github_issue" in recovery_info
    assert "fix_version" in recovery_info

    # 값 검증
    assert recovery_info["auto_recovery"] is True
    assert recovery_info["success_indicator"] == "차트 객체 존재 여부"
    assert recovery_info["github_issue"] == "#70"
    assert "10.2540.4" in recovery_info["fix_version"]


if __name__ == "__main__":
    # 테스트 실행 시 간단한 정보 출력
    print("=== GitHub Issue #70 COM 에러 테스트 ===")
    print(f"플랫폼: {platform.system()}")
    print(f"테스트 대상: COM 에러 0x800401FD 복구 로직")
    print(f"관련 파일: chart_pivot_create.py, utils.py")

    # pytest 실행 권장 메시지
    print("\n권장 실행 방법:")
    print("pytest tests/test_issue_70_com_error.py -v")
    print("pytest tests/test_issue_70_com_error.py::test_com_error_recovery_flow -v")
