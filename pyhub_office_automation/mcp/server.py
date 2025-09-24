"""
PyHub Office Automation MCP Server

Excel/HWP 자동화 기능을 AI 서비스와 연동하기 위한 FastMCP 서버
분석에 필수적인 최소한의 도구만 제공하여 빠른 로딩과 안정성 확보
"""

import json
import logging
from typing import Any, Dict, Optional

from fastmcp import FastMCP
from fastmcp.exceptions import ToolError

from pyhub_office_automation.excel.chart_list import chart_list
from pyhub_office_automation.excel.data_analyze import data_analyze
from pyhub_office_automation.excel.range_read import range_read
from pyhub_office_automation.excel.table_read import table_read
from pyhub_office_automation.excel.workbook_info import workbook_info

# 기존 Excel 자동화 함수 import
from pyhub_office_automation.excel.workbook_list import workbook_list
from pyhub_office_automation.version import get_version

# 로깅 설정
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# FastMCP 서버 인스턴스 생성
mcp = FastMCP(
    name="PyHub Office Automation MCP", version=get_version(), instructions="Excel 분석을 위한 최소 도구 세트 (FastMCP 기반)"
)

# =============================================================================
# Resources - 상태 정보 제공
# =============================================================================


@mcp.resource("resource://excel/workbooks")
def get_workbooks() -> Dict[str, Any]:
    """현재 열린 Excel 워크북 목록 제공"""
    try:
        # 기존 workbook_list 함수 재사용하여 JSON 형태로 반환
        result = workbook_list(detailed=True, output_format="json")

        # 문자열 결과를 JSON으로 파싱
        if isinstance(result, str):
            return json.loads(result)
        return result

    except Exception as e:
        logger.error(f"get_workbooks error: {e}")
        return {"error": "워크북 목록을 가져오는 중 오류 발생", "workbooks": [], "total": 0}


@mcp.resource("resource://excel/workbook/{name}/info")
def get_workbook_info(name: str) -> Dict[str, Any]:
    """특정 워크북의 상세 정보 제공"""
    try:
        # 워크북명으로 상세 정보 조회
        result = workbook_info(
            workbook_name=name, include_sheets=True, include_charts=True, include_pivot_tables=True, output_format="json"
        )

        if isinstance(result, str):
            return json.loads(result)
        return result

    except FileNotFoundError:
        return {"error": f"워크북 '{name}'을 찾을 수 없습니다"}
    except Exception as e:
        logger.error(f"get_workbook_info error for {name}: {e}")
        return {"error": "워크북 정보를 가져오는 중 오류 발생"}


# =============================================================================
# Tools - 핵심 분석 기능
# =============================================================================


@mcp.tool
def excel_workbook_info(workbook_name: Optional[str] = None) -> Dict[str, Any]:
    """
    워크북 구조 분석 및 상세 정보 제공

    Args:
        workbook_name: 분석할 워크북명 (미지정시 활성 워크북)

    Returns:
        워크북 구조, 시트 목록, 차트/피벗테이블 정보
    """
    try:
        result = workbook_info(
            workbook_name=workbook_name,
            include_sheets=True,
            include_charts=True,
            include_pivot_tables=True,
            include_properties=True,
            output_format="json",
        )

        if isinstance(result, str):
            return json.loads(result)
        return result

    except FileNotFoundError:
        raise ToolError(f"워크북을 찾을 수 없습니다: {workbook_name or '활성 워크북'}")
    except Exception as e:
        logger.error(f"excel_workbook_info error: {e}")
        raise ToolError("워크북 정보 분석 중 오류가 발생했습니다")


@mcp.tool
def excel_range_read(range_address: str, workbook_name: Optional[str] = None, sheet: Optional[str] = None) -> Dict[str, Any]:
    """
    Excel 셀 범위 데이터 읽기

    Args:
        range_address: 읽을 셀 범위 (예: "A1:C10")
        workbook_name: 대상 워크북명 (미지정시 활성 워크북)
        sheet: 대상 시트명 (미지정시 활성 시트)

    Returns:
        범위 데이터와 메타정보
    """
    try:
        result = range_read(range_address=range_address, workbook_name=workbook_name, sheet=sheet, output_format="json")

        if isinstance(result, str):
            return json.loads(result)
        return result

    except FileNotFoundError:
        raise ToolError(f"워크북을 찾을 수 없습니다: {workbook_name or '활성 워크북'}")
    except ValueError as e:
        raise ToolError(f"잘못된 범위 주소입니다: {range_address}")
    except Exception as e:
        logger.error(f"excel_range_read error: {e}")
        raise ToolError("범위 읽기 중 오류가 발생했습니다")


@mcp.tool
def excel_table_read(
    workbook_name: Optional[str] = None, sheet: Optional[str] = None, table_name: Optional[str] = None
) -> Dict[str, Any]:
    """
    Excel 테이블 데이터를 DataFrame으로 읽기

    Args:
        workbook_name: 대상 워크북명 (미지정시 활성 워크북)
        sheet: 대상 시트명 (미지정시 활성 시트)
        table_name: 테이블명 (미지정시 첫 번째 테이블)

    Returns:
        테이블 데이터와 구조 정보
    """
    try:
        result = table_read(workbook_name=workbook_name, sheet=sheet, table_name=table_name, output_format="json")

        if isinstance(result, str):
            return json.loads(result)
        return result

    except FileNotFoundError:
        raise ToolError(f"워크북을 찾을 수 없습니다: {workbook_name or '활성 워크북'}")
    except ValueError as e:
        raise ToolError(f"테이블을 찾을 수 없습니다: {table_name or '기본 테이블'}")
    except Exception as e:
        logger.error(f"excel_table_read error: {e}")
        raise ToolError("테이블 읽기 중 오류가 발생했습니다")


@mcp.tool
def excel_data_analyze(
    workbook_name: Optional[str] = None, sheet: Optional[str] = None, data_range: Optional[str] = None
) -> Dict[str, Any]:
    """
    데이터 구조 자동 분석 및 피벗테이블 추천

    Args:
        workbook_name: 분석할 워크북명 (미지정시 활성 워크북)
        sheet: 분석할 시트명 (미지정시 활성 시트)
        data_range: 분석할 데이터 범위 (미지정시 자동 감지)

    Returns:
        데이터 구조 분석 결과 및 추천사항
    """
    try:
        result = data_analyze(workbook_name=workbook_name, sheet=sheet, data_range=data_range, output_format="json")

        if isinstance(result, str):
            return json.loads(result)
        return result

    except FileNotFoundError:
        raise ToolError(f"워크북을 찾을 수 없습니다: {workbook_name or '활성 워크북'}")
    except Exception as e:
        logger.error(f"excel_data_analyze error: {e}")
        raise ToolError("데이터 분석 중 오류가 발생했습니다")


@mcp.tool
def excel_chart_list(workbook_name: Optional[str] = None, sheet: Optional[str] = None) -> Dict[str, Any]:
    """
    워크시트의 차트 목록 및 정보 조회

    Args:
        workbook_name: 대상 워크북명 (미지정시 활성 워크북)
        sheet: 대상 시트명 (미지정시 활성 시트)

    Returns:
        차트 목록과 각 차트의 상세 정보
    """
    try:
        result = chart_list(workbook_name=workbook_name, sheet=sheet, output_format="json")

        if isinstance(result, str):
            return json.loads(result)
        return result

    except FileNotFoundError:
        raise ToolError(f"워크북을 찾을 수 없습니다: {workbook_name or '활성 워크북'}")
    except Exception as e:
        logger.error(f"excel_chart_list error: {e}")
        raise ToolError("차트 목록 조회 중 오류가 발생했습니다")


# =============================================================================
# 서버 인스턴스 내보내기
# =============================================================================

# MCP 서버 인스턴스를 외부에서 사용할 수 있도록 내보내기
__all__ = ["mcp"]

if __name__ == "__main__":
    # 개발용 테스트 실행
    print("PyHub Office Automation MCP Server")
    print(f"Name: {mcp.name}")
    print(f"Version: {mcp.version}")
    print(f"Instructions: {mcp.instructions}")

    # 개발 모드에서는 간단히 표시
    print(f"\n=== Available Resources (2) ===")
    print("- resource://excel/workbooks")
    print("- resource://excel/workbook/{name}/info")

    print(f"\n=== Available Tools (5) ===")
    expected_tools = ["excel_workbook_info", "excel_range_read", "excel_table_read", "excel_data_analyze", "excel_chart_list"]
    for tool_name in expected_tools:
        print(f"- {tool_name}")
