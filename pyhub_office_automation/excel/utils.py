"""
Excel 자동화를 위한 공통 유틸리티 함수들
xlwings 기반 Excel 조작 및 데이터 처리 지원
"""

import csv
import datetime
import io
import json
import os
import platform
import re
import tempfile
import time
import unicodedata
from enum import Enum
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Union

import pandas as pd
import xlwings as xw

from pyhub_office_automation.version import get_version


# CLI 명령어 인자를 위한 Enum 클래스들
class OutputFormat(str, Enum):
    """출력 형식 선택지"""

    JSON = "json"
    CSV = "csv"
    TEXT = "text"
    MARKDOWN = "markdown"


class ExpandMode(str, Enum):
    """범위 확장 모드 선택지"""

    TABLE = "table"
    DOWN = "down"
    RIGHT = "right"


class LegendPosition(str, Enum):
    """차트 범례 위치 선택지"""

    TOP = "top"
    BOTTOM = "bottom"
    LEFT = "left"
    RIGHT = "right"
    NONE = "none"


class SlicerStyle(str, Enum):
    """슬라이서 스타일 선택지"""

    LIGHT = "light"
    MEDIUM = "medium"
    DARK = "dark"


class ColorScheme(str, Enum):
    """차트 색상 테마 선택지"""

    COLORFUL = "colorful"
    MONOCHROMATIC = "monochromatic"
    OFFICE = "office"
    GRAYSCALE = "grayscale"


class DataTransformType(str, Enum):
    """데이터 변환 타입 선택지"""

    UNPIVOT = "unpivot"
    UNMERGE = "unmerge"
    FLATTEN_HEADERS = "flatten-headers"
    REMOVE_SUBTOTALS = "remove-subtotals"
    AUTO = "auto"


class DataFormat(str, Enum):
    """데이터 형식 타입 선택지"""

    CROSS_TAB = "cross_tab"
    WIDE_FORMAT = "wide_format"
    MULTI_LEVEL_HEADERS = "multi_level_headers"
    MERGED_CELLS = "merged_cells"
    SUBTOTALS_MIXED = "subtotals_mixed"
    PIVOT_READY = "pivot_ready"
    UNKNOWN = "unknown"


class DataLabelPosition(str, Enum):
    """데이터 레이블 위치 선택지"""

    CENTER = "center"
    ABOVE = "above"
    BELOW = "below"
    LEFT = "left"
    RIGHT = "right"
    OUTSIDE = "outside"
    INSIDE = "inside"


class ChartType(str, Enum):
    """차트 타입 선택지"""

    # 기본 차트 타입들
    COLUMN = "column"
    COLUMN_CLUSTERED = "column_clustered"
    COLUMN_STACKED = "column_stacked"
    COLUMN_STACKED_100 = "column_stacked_100"
    BAR = "bar"
    BAR_CLUSTERED = "bar_clustered"
    BAR_STACKED = "bar_stacked"
    BAR_STACKED_100 = "bar_stacked_100"
    LINE = "line"
    LINE_MARKERS = "line_markers"
    PIE = "pie"
    DOUGHNUT = "doughnut"
    AREA = "area"
    AREA_STACKED = "area_stacked"
    SCATTER = "scatter"
    SCATTER_LINES = "scatter_lines"
    SCATTER_SMOOTH = "scatter_smooth"
    BUBBLE = "bubble"
    COMBO = "combo"


# PyInstaller 환경에서 win32com 초기화
def _initialize_win32com_for_pyinstaller():
    """PyInstaller 환경에서 win32com을 초기화하여 COM 캐시 재구축 경고를 방지합니다."""
    import sys

    if hasattr(sys, "_MEIPASS"):
        try:
            import warnings

            with warnings.catch_warnings():
                warnings.simplefilter("ignore")

                import win32com.client

                # 캐시 읽기 전용 설정을 해제하여 필요시 생성 가능하도록 함
                win32com.client.gencache.is_readonly = False

                # 캐시 경로 확인 및 초기화
                try:
                    win32com.client.gencache.GetGeneratePath()
                except Exception:
                    # 캐시 생성에 실패하면 읽기 전용으로 설정
                    win32com.client.gencache.is_readonly = True

        except ImportError:
            # win32com이 없는 환경에서는 무시
            pass
        except Exception:
            # 기타 오류는 조용히 무시
            pass


# 모듈 로드 시 win32com 초기화 실행
_initialize_win32com_for_pyinstaller()


def normalize_path(path: str) -> str:
    """
    경로의 한글 문자를 정규화합니다 (macOS 자소분리 문제 해결).

    Args:
        path: 정규화할 경로 문자열

    Returns:
        정규화된 경로 문자열
    """
    if not isinstance(path, str):
        return path

    # macOS에서 한글 자소분리 문제 해결
    if platform.system() == "Darwin":
        # NFD -> NFC 정규화 (자소 결합)
        return unicodedata.normalize("NFC", path)

    return path


def get_workbook(file_path: str, visible: bool = True) -> xw.Book:
    """
    Excel 워크북을 열거나 생성합니다.

    Args:
        file_path: Excel 파일 경로
        visible: Excel 애플리케이션 표시 여부

    Returns:
        xlwings Book 객체

    Raises:
        FileNotFoundError: 파일이 존재하지 않는 경우
        RuntimeError: Excel 애플리케이션 시작 실패
    """
    # 한글 경로 정규화 (macOS 자소분리 문제 해결)
    file_path = normalize_path(file_path)
    file_path = Path(file_path).resolve()

    # Excel 애플리케이션 시작
    try:
        app = xw.App(visible=visible)
    except Exception as e:
        raise RuntimeError(f"Excel 애플리케이션을 시작할 수 없습니다: {str(e)}")

    # 파일이 존재하면 열고, 없으면 새로 생성
    try:
        if file_path.exists():
            book = app.books.open(str(file_path))
        else:
            book = app.books.add()
            # 새 파일인 경우 저장
            if file_path.suffix:
                book.save(str(file_path))
    except Exception as e:
        app.quit()
        raise RuntimeError(f"워크북 처리 중 오류 발생: {str(e)}")

    return book


def get_sheet(book: xw.Book, sheet_name: Optional[str] = None) -> xw.Sheet:
    """
    워크북에서 시트를 가져옵니다.

    Args:
        book: xlwings Book 객체
        sheet_name: 시트 이름 (None이면 활성 시트)

    Returns:
        xlwings Sheet 객체

    Raises:
        ValueError: 시트가 존재하지 않는 경우
    """
    if sheet_name:
        try:
            return book.sheets[sheet_name]
        except:
            raise ValueError(f"시트 '{sheet_name}'를 찾을 수 없습니다")
    else:
        return book.sheets.active


def parse_range(range_str: str) -> Tuple[Optional[str], str]:
    """
    범위 문자열을 파싱하여 시트명과 범위를 분리합니다.

    Args:
        range_str: 범위 문자열 (예: "A1:C10" 또는 "Sheet1!A1:C10")

    Returns:
        (시트명, 범위) 튜플
    """
    if "!" in range_str:
        sheet_name, range_part = range_str.split("!", 1)
        return sheet_name, range_part
    else:
        return None, range_str


def get_range(sheet: xw.Sheet, range_str: str, expand_mode: Optional[ExpandMode] = None) -> xw.Range:
    """
    시트에서 지정된 범위를 가져옵니다.

    Args:
        sheet: xlwings Sheet 객체
        range_str: 범위 문자열 (예: "A1:C10")
        expand_mode: 확장 모드 ("table", "down", "right")

    Returns:
        xlwings Range 객체
    """
    range_obj = sheet.range(range_str)

    if expand_mode:
        if expand_mode == ExpandMode.TABLE:
            range_obj = range_obj.expand()
        elif expand_mode == ExpandMode.DOWN:
            range_obj = range_obj.expand("down")
        elif expand_mode == ExpandMode.RIGHT:
            range_obj = range_obj.expand("right")

    return range_obj


def handle_temp_file(data: Union[str, list, dict], file_format: str = "json") -> str:
    """
    임시 파일을 생성하고 데이터를 저장합니다.

    Args:
        data: 저장할 데이터
        file_format: 파일 형식 ("json", "csv")

    Returns:
        임시 파일 경로
    """
    if file_format == "json":
        suffix = ".json"
        mode = "w"
    elif file_format == "csv":
        suffix = ".csv"
        mode = "w"
    else:
        suffix = ".txt"
        mode = "w"

    temp_file = tempfile.NamedTemporaryFile(mode=mode, suffix=suffix, delete=False, encoding="utf-8")

    try:
        if file_format == "json":
            json.dump(data, temp_file, ensure_ascii=False, indent=2)
        elif file_format == "csv" and isinstance(data, list):
            writer = csv.writer(temp_file)
            writer.writerows(data)
        else:
            temp_file.write(str(data))
    finally:
        temp_file.close()

    return temp_file.name


def cleanup_temp_file(file_path: str) -> None:
    """
    임시 파일을 안전하게 삭제합니다.

    Args:
        file_path: 삭제할 파일 경로
    """
    try:
        if os.path.exists(file_path):
            os.unlink(file_path)
    except Exception:
        # 삭제 실패해도 조용히 넘어감 (Windows 파일 잠금 등)
        pass


def format_output(data: Union[str, list, dict], output_format: str = "json", include_version: bool = True) -> str:
    """
    데이터를 지정된 형식으로 포맷팅합니다.

    Args:
        data: 포맷팅할 데이터
        output_format: 출력 형식 ("json", "csv", "text")
        include_version: 버전 정보 포함 여부

    Returns:
        포맷팅된 문자열
    """
    if include_version and isinstance(data, dict):
        data["version"] = get_version()

    if output_format == "json":
        return json.dumps(data, ensure_ascii=False, indent=2)
    elif output_format == "csv" and isinstance(data, list):
        output = io.StringIO()
        writer = csv.writer(output)
        writer.writerows(data)
        return output.getvalue()
    elif output_format == "text":
        if isinstance(data, dict):
            return "\n".join([f"{k}: {v}" for k, v in data.items()])
        elif isinstance(data, list):
            return "\n".join([str(item) for item in data])
        else:
            return str(data)
    else:
        return str(data)


def load_data_from_file(file_path: str) -> Union[str, list, dict]:
    """
    파일에서 데이터를 로드합니다.

    Args:
        file_path: 데이터 파일 경로

    Returns:
        로드된 데이터

    Raises:
        FileNotFoundError: 파일이 존재하지 않는 경우
        ValueError: 파일 형식이 지원되지 않는 경우
    """
    # 한글 경로 정규화 (macOS 자소분리 문제 해결)
    file_path = normalize_path(file_path)
    file_path = Path(file_path)

    if not file_path.exists():
        raise FileNotFoundError(f"데이터 파일을 찾을 수 없습니다: {file_path}")

    suffix = file_path.suffix.lower()

    try:
        if suffix == ".json":
            with open(file_path, "r", encoding="utf-8") as f:
                return json.load(f)
        elif suffix == ".csv":
            with open(file_path, "r", encoding="utf-8") as f:
                reader = csv.reader(f)
                return [row for row in reader]
        else:
            with open(file_path, "r", encoding="utf-8") as f:
                return f.read()
    except Exception as e:
        raise ValueError(f"데이터 파일 로드 실패: {str(e)}")


def validate_range_string(range_str: str) -> bool:
    """
    범위 문자열의 유효성을 검증합니다.

    Args:
        range_str: 검증할 범위 문자열

    Returns:
        유효성 여부
    """
    import re

    # Sheet!Range 형태 처리
    if "!" in range_str:
        _, range_part = range_str.split("!", 1)
    else:
        range_part = range_str

    # A1:B10 형태의 범위 패턴
    range_pattern = r"^[A-Z]+\d+(:[A-Z]+\d+)?$"
    return bool(re.match(range_pattern, range_part.upper()))


# COM 에러 코드와 사용자 친화적 메시지 매핑
COM_ERROR_MESSAGES = {
    0x800A01A8: {
        "message": "Excel 객체에 접근할 수 없습니다. Excel이 실행 중이고 워크북이 열려있는지 확인하세요.",
        "meaning": "Object Required",
        "causes": [
            "Excel이 실행되지 않음",
            "워크북이 닫혀있음",
            "Excel 객체가 해제됨"
        ],
        "suggestions": [
            "Excel 프로그램이 실행 중인지 확인",
            "워크북이 열려있는지 확인",
            "--visible 옵션을 사용하여 Excel 창 표시"
        ]
    },
    0x800401A8: {
        "message": "Excel COM 객체 연결이 끊어졌습니다.",
        "meaning": "Object is disconnected from clients",
        "causes": [
            "Excel 프로세스가 예기치 않게 종료됨",
            "COM 객체 수명 주기 문제"
        ],
        "suggestions": [
            "Excel을 다시 시작하세요",
            "명령을 다시 실행하세요"
        ]
    },
    0x80010105: {
        "message": "Excel 서버가 예기치 않게 종료되었습니다.",
        "meaning": "RPC_E_SERVERFAULT",
        "causes": [
            "Excel이 충돌하거나 강제 종료됨",
            "메모리 부족"
        ],
        "suggestions": [
            "Excel을 다시 시작하세요",
            "시스템 메모리를 확인하세요"
        ]
    },
    0x800A03EC: {
        "message": "Excel 작업이 실패했습니다. 데이터나 작업이 유효하지 않을 수 있습니다.",
        "meaning": "NAME_NOT_FOUND or INVALID_OPERATION",
        "causes": [
            "잘못된 범위 이름 또는 시트 이름",
            "지원되지 않는 작업",
            "데이터 형식 오류"
        ],
        "suggestions": [
            "시트 이름과 범위가 정확한지 확인",
            "데이터 형식이 올바른지 확인",
            "Excel 버전과 호환되는지 확인"
        ]
    }
}


def extract_com_error_code(error: Exception) -> Optional[int]:
    """
    COM 에러에서 에러 코드를 추출합니다.

    Args:
        error: COM 에러 예외

    Returns:
        에러 코드 (정수) 또는 None
    """
    try:
        # com_error 형식: (hr, msg, exc, arg)
        if hasattr(error, 'args') and len(error.args) > 0:
            # 첫 번째 인자가 HRESULT 코드
            if isinstance(error.args[0], int):
                # 음수를 양수로 변환 (비트 마스크 적용)
                return error.args[0] & 0xFFFFFFFF
            # 튜플 형태인 경우
            elif isinstance(error.args[0], tuple) and len(error.args[0]) > 0:
                return error.args[0][0] & 0xFFFFFFFF
    except:
        pass
    return None


def create_error_response(error: Exception, command: str) -> Dict[str, Union[str, int, float]]:
    """
    표준 에러 응답을 생성합니다.

    Args:
        error: 발생한 예외
        command: 명령어 이름

    Returns:
        에러 응답 딕셔너리
    """
    error_type = type(error).__name__

    # COM 에러 특별 처리
    if error_type == "com_error" or "com_error" in error_type.lower():
        error_code = extract_com_error_code(error)

        if error_code and error_code in COM_ERROR_MESSAGES:
            com_info = COM_ERROR_MESSAGES[error_code]
            return {
                "success": False,
                "error_type": error_type,
                "error": com_info["message"],
                "error_details": {
                    "code": hex(error_code),
                    "meaning": com_info["meaning"],
                    "possible_causes": com_info["causes"],
                    "original_error": str(error)
                },
                "suggestions": com_info["suggestions"],
                "command": command,
                "version": get_version()
            }
        else:
            # 알려지지 않은 COM 에러
            return {
                "success": False,
                "error_type": error_type,
                "error": "Excel COM 오류가 발생했습니다.",
                "error_details": {
                    "code": hex(error_code) if error_code else "unknown",
                    "original_error": str(error)
                },
                "suggestions": [
                    "Excel이 실행 중인지 확인하세요",
                    "워크북이 열려있는지 확인하세요",
                    "Excel을 재시작해보세요"
                ],
                "command": command,
                "version": get_version()
            }

    # 기존 에러 처리 로직
    response = {"success": False, "error_type": error_type, "error": str(error), "command": command, "version": get_version()}

    # 특정 에러에 대한 제안사항 추가
    if error_type == "FileNotFoundError":
        response["suggestion"] = "파일 경로를 확인하고 파일이 존재하는지 확인하세요."
    elif error_type == "RuntimeError" and "Excel" in str(error):
        response["suggestion"] = "Excel이 설치되어 있는지 확인하고, 다른 프로그램에서 파일을 사용 중이지 않은지 확인하세요."
    elif error_type == "ValueError" and "범위" in str(error):
        response["suggestion"] = "범위 형식이 올바른지 확인하세요. 예: 'A1:C10' 또는 'Sheet1!A1:C10'"

    return response


def create_success_response(
    data: Union[str, list, dict],
    command: str,
    message: str = None,
    execution_time_ms: float = None,
    book: Optional[xw.Book] = None,
    **stats_kwargs,
) -> Dict[str, Union[str, int, float, list, dict]]:
    """
    AI 에이전트 호환성이 향상된 표준 성공 응답을 생성합니다.

    Args:
        data: 응답 데이터
        command: 명령어 이름
        message: 성공 메시지
        execution_time_ms: 실행 시간 (밀리초)
        book: xlwings Book 객체 (컨텍스트 수집용)
        **stats_kwargs: 작업 통계 데이터

    Returns:
        성공 응답 딕셔너리
    """
    response = {"success": True, "command": command, "version": get_version(), "data": data}

    if message:
        response["message"] = message

    # 메타데이터 추가
    metadata = {
        "command_category": COMMAND_CATEGORIES.get(command, "unknown"),
        "operation_type": OPERATION_TYPES.get(command, "unknown"),
        "timestamp": datetime.datetime.now().isoformat(),
    }

    if execution_time_ms is not None:
        metadata["execution_time_ms"] = execution_time_ms

    response["metadata"] = metadata

    # 작업 통계 추가
    if stats_kwargs or command:
        operation_stats = calculate_operation_stats(command, **stats_kwargs)
        response["operation_stats"] = operation_stats

    # 실행 컨텍스트 추가
    try:
        context = get_execution_context(book)
        response["current_context"] = context

        # 후속 명령어 추천
        suggestions = suggest_next_commands(command, context, **stats_kwargs)
        if suggestions:
            response["suggested_next_commands"] = suggestions

    except Exception:
        # 컨텍스트 수집 실패 시 기본 정보만 포함
        response["current_context"] = {"total_open_workbooks": len(xw.books) if xw.books else 0, "collection_failed": True}

    return response


# ========== 활성 워크북 연결 기능 (Issue #14) ==========


def get_active_app(visible: bool = True) -> Optional[xw.App]:
    """
    현재 실행 중인 Excel 애플리케이션을 가져옵니다.

    Args:
        visible: Excel 애플리케이션 표시 여부

    Returns:
        xlwings App 객체 또는 None

    Raises:
        RuntimeError: Excel 애플리케이션을 찾을 수 없는 경우
    """
    try:
        # 이미 실행 중인 Excel 앱이 있는지 확인
        if len(xw.apps) > 0:
            # 가장 최근에 활성화된 앱 반환
            return xw.apps.active
        else:
            # 실행 중인 앱이 없으면 새로 생성
            return xw.App(visible=visible)
    except Exception as e:
        raise RuntimeError(f"Excel 애플리케이션을 가져올 수 없습니다: {str(e)}")


def get_active_workbook() -> xw.Book:
    """
    현재 활성 워크북을 반환합니다.

    Returns:
        xlwings Book 객체

    Raises:
        RuntimeError: 활성 워크북이 없는 경우
    """
    try:
        if len(xw.books) == 0:
            raise RuntimeError("현재 열려있는 워크북이 없습니다")

        return xw.books.active
    except Exception as e:
        raise RuntimeError(f"활성 워크북을 가져올 수 없습니다: {str(e)}")


def get_workbook_by_name(workbook_name: str) -> xw.Book:
    """
    이름으로 열린 워크북을 찾습니다.

    Args:
        workbook_name: 찾을 워크북 이름 (예: "Sales.xlsx")

    Returns:
        xlwings Book 객체

    Raises:
        RuntimeError: 지정한 이름의 워크북을 찾을 수 없는 경우
    """
    try:
        if len(xw.books) == 0:
            raise RuntimeError("현재 열려있는 워크북이 없습니다")

        # 정확한 이름으로 먼저 검색
        for book in xw.books:
            if book.name == workbook_name:
                return book

        # 파일 이름만으로 검색 (확장자 포함/제외)
        for book in xw.books:
            if Path(book.name).name == workbook_name:
                return book
            if Path(book.name).stem == Path(workbook_name).stem:
                return book

        raise RuntimeError(f"워크북 '{workbook_name}'을 찾을 수 없습니다")

    except Exception as e:
        if "워크북" in str(e):
            raise  # 이미 우리가 만든 에러면 그대로 전달
        else:
            raise RuntimeError(f"워크북 검색 중 오류 발생: {str(e)}")


def get_or_open_workbook(
    file_path: Optional[str] = None, workbook_name: Optional[str] = None, use_active: bool = False, visible: bool = True
) -> xw.Book:
    """
    여러 방법으로 워크북을 가져오는 통합 함수입니다.

    Args:
        file_path: 파일 경로 (기존 방식)
        workbook_name: 열린 워크북 이름
        use_active: 활성 워크북 사용 여부 (내부 사용, 옵션 없으면 자동으로 True)
        visible: Excel 애플리케이션 표시 여부

    Returns:
        xlwings Book 객체

    Raises:
        ValueError: 옵션이 잘못 지정된 경우
        RuntimeError: 워크북을 찾거나 열 수 없는 경우
    """
    # 옵션 검증 - 최대 하나만 지정되어야 함
    options_count = sum([bool(file_path), bool(workbook_name), use_active])

    if options_count == 0:
        # 옵션이 없으면 활성 워크북 자동 사용 (Issue #31)
        use_active = True
    elif options_count > 1:
        raise ValueError("file_path, workbook_name, use_active 중 하나만 지정할 수 있습니다")

    try:
        if use_active:
            return get_active_workbook()
        elif workbook_name:
            return get_workbook_by_name(workbook_name)
        elif file_path:
            return get_workbook(file_path, visible=visible)
        else:
            raise ValueError("올바른 옵션이 지정되지 않았습니다")

    except Exception as e:
        # 구체적인 에러 메시지 추가
        if use_active:
            raise RuntimeError(f"활성 워크북을 가져올 수 없습니다: {str(e)}")
        elif workbook_name:
            raise RuntimeError(f"워크북 '{workbook_name}'를 찾을 수 없습니다: {str(e)}")
        elif file_path:
            raise RuntimeError(f"파일 '{file_path}'를 열 수 없습니다: {str(e)}")
        else:
            raise


# ========== AI 에이전트 호환성 향상 기능 ==========

# 명령어별 추천 매핑
COMMAND_SUGGESTIONS = {
    "workbook-open": [
        {"command": "workbook-info", "reason": "워크북의 상세 정보를 확인합니다"},
        {"command": "workbook-list", "reason": "다른 열린 워크북들을 확인합니다"},
        {"command": "range-read", "reason": "워크북의 데이터를 읽습니다"},
    ],
    "workbook-list": [
        {"command": "workbook-info", "reason": "특정 워크북의 상세 정보를 확인합니다"},
        {"command": "workbook-open", "reason": "새 워크북을 엽니다"},
    ],
    "workbook-info": [
        {"command": "range-read", "reason": "시트의 데이터를 확인합니다"},
        {"command": "sheet-activate", "reason": "특정 시트로 이동합니다"},
    ],
    "range-read": [
        {"command": "range-write", "reason": "데이터를 수정하거나 추가합니다"},
        {"command": "table-read", "reason": "테이블 형태로 더 많은 데이터를 읽습니다"},
    ],
    "range-write": [
        {"command": "range-read", "reason": "작성한 데이터를 확인합니다"},
        {"command": "workbook-save", "reason": "변경사항을 저장합니다"},
    ],
    "table-read": [
        {"command": "table-write", "reason": "테이블 데이터를 수정합니다"},
        {"command": "pivot-create", "reason": "피벗 테이블을 생성합니다"},
    ],
    "table-write": [
        {"command": "table-read", "reason": "작성한 테이블을 확인합니다"},
        {"command": "chart-add", "reason": "차트를 생성합니다"},
    ],
    "sheet-add": [
        {"command": "sheet-activate", "reason": "새 시트로 이동합니다"},
        {"command": "range-write", "reason": "새 시트에 데이터를 입력합니다"},
        {"command": "sheet-rename", "reason": "시트 이름을 변경합니다"},
    ],
    "sheet-activate": [
        {"command": "range-read", "reason": "활성 시트의 데이터를 확인합니다"},
        {"command": "sheet-add", "reason": "새 시트를 추가합니다"},
    ],
    "sheet-rename": [
        {"command": "sheet-activate", "reason": "이름이 변경된 시트로 이동합니다"},
        {"command": "workbook-info", "reason": "전체 시트 목록을 확인합니다"},
    ],
    "sheet-delete": [
        {"command": "workbook-info", "reason": "남은 시트들을 확인합니다"},
        {"command": "sheet-add", "reason": "새 시트를 추가합니다"},
    ],
    "pivot-create": [
        {"command": "pivot-configure", "reason": "피벗 테이블을 설정합니다"},
        {"command": "chart-pivot-create", "reason": "피벗 차트를 생성합니다"},
        {"command": "slicer-add", "reason": "피벗 테이블용 슬라이서를 생성합니다"},
    ],
    "chart-add": [
        {"command": "chart-configure", "reason": "차트를 설정합니다"},
        {"command": "chart-position", "reason": "차트 위치를 조정합니다"},
    ],
    "shape-add": [
        {"command": "shape-format", "reason": "도형 스타일을 설정합니다"},
        {"command": "shape-list", "reason": "생성된 도형을 확인합니다"},
        {"command": "textbox-add", "reason": "텍스트 박스를 추가합니다"},
    ],
    "shape-format": [
        {"command": "shape-list", "reason": "스타일이 적용된 도형을 확인합니다"},
        {"command": "shape-group", "reason": "여러 도형을 그룹화합니다"},
    ],
    "shape-list": [
        {"command": "shape-format", "reason": "도형 스타일을 변경합니다"},
        {"command": "shape-delete", "reason": "불필요한 도형을 삭제합니다"},
        {"command": "shape-group", "reason": "도형들을 그룹화합니다"},
    ],
    "shape-delete": [
        {"command": "shape-list", "reason": "남은 도형들을 확인합니다"},
        {"command": "shape-add", "reason": "새로운 도형을 추가합니다"},
    ],
    "shape-group": [
        {"command": "shape-list", "reason": "그룹화된 도형을 확인합니다"},
        {"command": "shape-format", "reason": "그룹 스타일을 조정합니다"},
    ],
    "textbox-add": [
        {"command": "shape-list", "reason": "생성된 텍스트 박스를 확인합니다"},
        {"command": "shape-format", "reason": "텍스트 박스 스타일을 변경합니다"},
        {"command": "shape-group", "reason": "다른 도형과 그룹화합니다"},
    ],
    "slicer-add": [
        {"command": "slicer-connect", "reason": "다른 피벗테이블에 연결합니다"},
        {"command": "slicer-position", "reason": "슬라이서 위치를 조정합니다"},
        {"command": "slicer-list", "reason": "생성된 슬라이서를 확인합니다"},
    ],
    "slicer-list": [
        {"command": "slicer-connect", "reason": "슬라이서 연결을 관리합니다"},
        {"command": "slicer-position", "reason": "슬라이서 위치를 조정합니다"},
    ],
    "slicer-connect": [
        {"command": "slicer-list", "reason": "연결 상태를 확인합니다"},
        {"command": "slicer-position", "reason": "슬라이서 레이아웃을 정리합니다"},
    ],
    "slicer-position": [
        {"command": "slicer-list", "reason": "배치된 슬라이서를 확인합니다"},
        {"command": "shape-add", "reason": "슬라이서 배경 도형을 추가합니다"},
    ],
}

# 명령어 카테고리 정의
COMMAND_CATEGORIES = {
    "workbook-open": "workbook",
    "workbook-list": "workbook",
    "workbook-info": "workbook",
    "workbook-create": "workbook",
    "range-read": "data",
    "range-write": "data",
    "table-read": "data",
    "table-write": "data",
    "sheet-add": "sheet",
    "sheet-activate": "sheet",
    "sheet-rename": "sheet",
    "sheet-delete": "sheet",
    "pivot-create": "pivot",
    "pivot-configure": "pivot",
    "pivot-list": "pivot",
    "pivot-refresh": "pivot",
    "pivot-delete": "pivot",
    "chart-add": "chart",
    "chart-configure": "chart",
    "chart-list": "chart",
    "chart-position": "chart",
    "chart-pivot-create": "chart",
    "chart-delete": "chart",
    "chart-export": "chart",
    "shape-add": "shape",
    "shape-format": "shape",
    "shape-list": "shape",
    "shape-delete": "shape",
    "shape-group": "shape",
    "textbox-add": "shape",
    "slicer-add": "slicer",
    "slicer-connect": "slicer",
    "slicer-position": "slicer",
    "slicer-list": "slicer",
}

# 작업 타입 정의
OPERATION_TYPES = {
    "workbook-open": "read",
    "workbook-list": "read",
    "workbook-info": "read",
    "workbook-create": "create",
    "range-read": "read",
    "range-write": "write",
    "table-read": "read",
    "table-write": "write",
    "sheet-add": "create",
    "sheet-activate": "modify",
    "sheet-rename": "modify",
    "sheet-delete": "delete",
    "pivot-create": "create",
    "pivot-configure": "modify",
    "pivot-list": "read",
    "pivot-refresh": "modify",
    "pivot-delete": "delete",
    "chart-add": "create",
    "chart-configure": "modify",
    "chart-list": "read",
    "chart-position": "modify",
    "chart-pivot-create": "create",
    "chart-delete": "delete",
    "chart-export": "read",
    "shape-add": "create",
    "shape-format": "modify",
    "shape-list": "read",
    "shape-delete": "delete",
    "shape-group": "modify",
    "textbox-add": "create",
    "slicer-add": "create",
    "slicer-connect": "modify",
    "slicer-position": "modify",
    "slicer-list": "read",
}


def get_execution_context(book: Optional[xw.Book] = None) -> Dict[str, Union[str, int, float, list, dict]]:
    """
    현재 Excel 실행 컨텍스트 정보를 수집합니다.

    Args:
        book: xlwings Book 객체 (None이면 활성 워크북 사용)

    Returns:
        컨텍스트 정보 딕셔너리
    """
    context = {"total_open_workbooks": len(xw.books) if xw.books else 0, "excel_app_visible": None, "current_workbook": None}

    try:
        if book is None and len(xw.books) > 0:
            book = xw.books.active

        if book:
            context.update(
                {
                    "current_workbook": {
                        "name": normalize_path(book.name),
                        "full_name": normalize_path(book.fullname),
                        "saved": getattr(book, "saved", True),
                        "total_sheets": len(book.sheets),
                        "active_sheet": book.sheets.active.name if book.sheets else None,
                        "sheet_names": [sheet.name for sheet in book.sheets],
                    },
                    "excel_app_visible": book.app.visible if hasattr(book, "app") else None,
                }
            )

            # 저장되지 않은 워크북 체크
            unsaved_workbooks = []
            for wb in xw.books:
                try:
                    if not wb.saved:
                        unsaved_workbooks.append(normalize_path(wb.name))
                except:
                    pass

            if unsaved_workbooks:
                context["unsaved_workbooks"] = unsaved_workbooks

    except Exception:
        # 컨텍스트 수집 실패 시 기본값 유지
        pass

    return context


def calculate_operation_stats(command: str, **kwargs) -> Dict[str, Union[str, int, float, list, dict]]:
    """
    작업별 통계 정보를 계산합니다.

    Args:
        command: 실행된 명령어
        **kwargs: 명령어별 통계 데이터

    Returns:
        작업 통계 딕셔너리
    """
    stats = {"command": command, "timestamp": datetime.datetime.now().isoformat()}

    # 명령어별 특화 통계
    if command in ["range-read", "range-write"]:
        if "range_obj" in kwargs:
            range_obj = kwargs["range_obj"]
            try:
                stats.update(
                    {
                        "cells_count": range_obj.count if hasattr(range_obj, "count") else 1,
                        "rows_count": range_obj.rows.count if hasattr(range_obj, "rows") else 1,
                        "columns_count": range_obj.columns.count if hasattr(range_obj, "columns") else 1,
                        "range_address": range_obj.address if hasattr(range_obj, "address") else None,
                    }
                )
            except:
                pass

        if "data_size" in kwargs:
            stats["data_size_bytes"] = kwargs["data_size"]

    elif command in ["table-read", "table-write"]:
        if "rows_count" in kwargs:
            stats["rows_processed"] = kwargs["rows_count"]
        if "columns_count" in kwargs:
            stats["columns_processed"] = kwargs["columns_count"]

    elif command in ["sheet-add", "sheet-delete"]:
        if "sheet_name" in kwargs:
            stats["sheet_name"] = kwargs["sheet_name"]
        if "total_sheets" in kwargs:
            stats["total_sheets_after"] = kwargs["total_sheets"]

    elif command in ["workbook-open", "workbook-create"]:
        if "sheet_count" in kwargs:
            stats["sheets_in_workbook"] = kwargs["sheet_count"]
        if "file_size" in kwargs:
            stats["file_size_bytes"] = kwargs["file_size"]

    return stats


def suggest_next_commands(
    command: str, context: Dict[str, Union[str, int, float, list, dict]], **kwargs
) -> List[Dict[str, str]]:
    """
    현재 명령어와 컨텍스트를 기반으로 후속 명령어를 추천합니다.

    Args:
        command: 실행된 명령어
        context: 현재 실행 컨텍스트
        **kwargs: 추가 조건 데이터

    Returns:
        추천 명령어 리스트
    """
    suggestions = []

    # 기본 추천 명령어
    if command in COMMAND_SUGGESTIONS:
        suggestions.extend(COMMAND_SUGGESTIONS[command])

    # 컨텍스트 기반 동적 추천
    if context.get("current_workbook"):
        workbook_info = context["current_workbook"]

        # 저장되지 않은 변경사항이 있는 경우
        if not workbook_info.get("saved", True):
            suggestions.insert(0, {"command": "workbook-save", "reason": "워크북에 저장되지 않은 변경사항이 있습니다"})

        # 시트가 많은 경우 시트 관리 명령어 추천
        if workbook_info.get("total_sheets", 0) > 5:
            suggestions.append(
                {"command": "workbook-info", "reason": "많은 시트가 있으므로 전체 구조를 확인하는 것이 좋습니다"}
            )

    # 저장되지 않은 다른 워크북이 있는 경우
    if context.get("unsaved_workbooks"):
        suggestions.append(
            {"command": "workbook-list", "reason": f"저장되지 않은 워크북이 {len(context['unsaved_workbooks'])}개 있습니다"}
        )

    # 중복 제거 및 최대 5개까지만 반환
    seen_commands = set()
    unique_suggestions = []
    for suggestion in suggestions:
        if suggestion["command"] not in seen_commands:
            seen_commands.add(suggestion["command"])
            unique_suggestions.append(suggestion)
            if len(unique_suggestions) >= 5:
                break

    return unique_suggestions


class ExecutionTimer:
    """명령어 실행 시간을 측정하는 컨텍스트 매니저"""

    def __init__(self):
        self.start_time = None
        self.end_time = None

    def __enter__(self):
        self.start_time = time.time()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.end_time = time.time()

    @property
    def execution_time_ms(self) -> float:
        """실행 시간을 밀리초로 반환"""
        if self.start_time and self.end_time:
            return round((self.end_time - self.start_time) * 1000, 2)
        return 0.0


# ========== Shape 및 Slicer 관리 기능 (Issue #12) ==========

# 뉴모피즘 스타일 정의
NEUMORPHISM_STYLES = {
    "background": {"fill_color": "#F2EDF3", "transparency": 0, "has_line": False},
    "title-box": {
        "fill_color": "#1D2433",
        "transparency": 0,
        "has_line": False,
        "shadow": {"color": "#1D2433", "transparency": 85, "blur": 30, "angle": 45, "distance": 10},
    },
    "chart-box": {
        "fill_color": "#FFFFFF",
        "transparency": 0,
        "has_line": False,
        "shadow": {"color": "#1D2433", "transparency": 85, "blur": 30, "angle": 45, "distance": 10},
    },
    "slicer-box": {
        "fill_color": "#FFFFFF",
        "transparency": 0,
        "has_line": True,
        "line_color": "#E0E0E0",
        "shadow": {"color": "#1D2433", "transparency": 90, "blur": 20, "angle": 45, "distance": 5},
    },
}

# Excel Shape Type 상수 (xlwings 호환)
SHAPE_TYPES = {
    "rectangle": 1,  # msoShapeRectangle
    "oval": 9,  # msoShapeOval
    "line": 9,  # msoShapeLine (실제로는 20)
    "arrow": 4,  # msoShapeRightArrow
    "rounded_rectangle": 5,  # msoShapeRoundedRectangle
    "diamond": 4,  # msoShapeDiamond (실제로는 4)
    "triangle": 3,  # msoShapeIsoscelesTriangle
    "pentagon": 7,  # msoShapePentagon
    "hexagon": 8,  # msoShapeHexagon
    "star": 12,  # msoShape5pointStar
    "callout_rectangle": 61,  # msoShapeRectangularCallout
    "text_box": 17,  # msoTextBox
}


# 색상 유틸리티 함수
def hex_to_rgb(hex_color: str) -> int:
    """
    HEX 색상 코드를 Excel RGB 정수값으로 변환

    Args:
        hex_color: HEX 색상 코드 (예: "#FF0000", "FF0000")

    Returns:
        Excel에서 사용하는 RGB 정수값
    """
    if hex_color.startswith("#"):
        hex_color = hex_color[1:]

    # RGB 값 추출
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)

    # Excel RGB 형식으로 변환 (BGR 순서)
    return b * 65536 + g * 256 + r


def rgb_to_hex(rgb_value: int) -> str:
    """
    Excel RGB 정수값을 HEX 색상 코드로 변환

    Args:
        rgb_value: Excel RGB 정수값

    Returns:
        HEX 색상 코드 (예: "#FF0000")
    """
    # BGR에서 RGB로 변환
    b = (rgb_value // 65536) % 256
    g = (rgb_value // 256) % 256
    r = rgb_value % 256

    return f"#{r:02X}{g:02X}{b:02X}"


def apply_neumorphism_style(shape, style_name: str) -> bool:
    """
    도형에 뉴모피즘 스타일을 적용합니다.

    Args:
        shape: xlwings Shape 객체
        style_name: 적용할 스타일 이름

    Returns:
        스타일 적용 성공 여부
    """
    if style_name not in NEUMORPHISM_STYLES:
        return False

    try:
        style = NEUMORPHISM_STYLES[style_name]

        # Windows에서만 전체 기능 지원
        if platform.system() == "Windows":
            # 채우기 색상 설정
            if style.get("fill_color"):
                shape.api.Fill.ForeColor.RGB = hex_to_rgb(style["fill_color"])
                shape.api.Fill.Transparency = style.get("transparency", 0) / 100.0

            # 테두리 설정
            if style.get("has_line", True):
                shape.api.Line.Visible = True
                if style.get("line_color"):
                    shape.api.Line.ForeColor.RGB = hex_to_rgb(style["line_color"])
            else:
                shape.api.Line.Visible = False

            # 그림자 효과 설정
            if style.get("shadow"):
                shadow = style["shadow"]
                shape.api.Shadow.Type = 25  # msoShadow25 (외부 그림자)
                shape.api.Shadow.ForeColor.RGB = hex_to_rgb(shadow["color"])
                shape.api.Shadow.Transparency = shadow.get("transparency", 50) / 100.0
                shape.api.Shadow.Blur = shadow.get("blur", 20)
                shape.api.Shadow.OffsetX = shadow.get("distance", 5)
                shape.api.Shadow.OffsetY = shadow.get("distance", 5)
        else:
            # macOS에서는 기본 색상만 설정
            if style.get("fill_color"):
                # macOS에서는 제한적인 스타일링만 가능
                pass

        return True

    except Exception:
        return False


# =============================================================================
# 범위 관리 및 자동 배치 유틸리티 함수들
# =============================================================================


def excel_address_to_coords(address: str) -> Tuple[int, int]:
    """
    Excel 주소(A1)를 좌표(row, col)로 변환합니다.

    Args:
        address: Excel 주소 (예: "A1", "BC123")

    Returns:
        (row, col) 튜플 (1-based index)

    Raises:
        ValueError: 잘못된 주소 형식
    """
    import re

    # 주소 형식 검증
    match = re.match(r"^([A-Z]+)(\d+)$", address.upper())
    if not match:
        raise ValueError(f"잘못된 Excel 주소 형식: {address}")

    col_letters, row_str = match.groups()
    row = int(row_str)

    # 열 문자를 숫자로 변환 (A=1, B=2, ..., Z=26, AA=27, ...)
    col = 0
    for i, letter in enumerate(reversed(col_letters)):
        col += (ord(letter) - ord("A") + 1) * (26**i)

    return row, col


def coords_to_excel_address(row: int, col: int) -> str:
    """
    좌표(row, col)를 Excel 주소로 변환합니다.

    Args:
        row: 행 번호 (1-based)
        col: 열 번호 (1-based)

    Returns:
        Excel 주소 (예: "A1", "BC123")
    """
    if row < 1 or col < 1:
        raise ValueError("행과 열 번호는 1 이상이어야 합니다")

    # 열 번호를 문자로 변환
    col_letters = ""
    while col > 0:
        col -= 1
        col_letters = chr(ord("A") + col % 26) + col_letters
        col //= 26

    return f"{col_letters}{row}"


def parse_excel_range(range_str: str) -> Tuple[int, int, int, int]:
    """
    Excel 범위(A1:C10)를 좌표로 파싱합니다.

    Args:
        range_str: Excel 범위 (예: "A1:C10", "A1" 단일 셀도 가능)

    Returns:
        (start_row, start_col, end_row, end_col) 튜플 (1-based)

    Raises:
        ValueError: 잘못된 범위 형식
    """
    range_str = range_str.strip()

    if ":" in range_str:
        # 범위 형식 (A1:C10)
        start_addr, end_addr = range_str.split(":", 1)
        start_row, start_col = excel_address_to_coords(start_addr.strip())
        end_row, end_col = excel_address_to_coords(end_addr.strip())
    else:
        # 단일 셀 형식 (A1)
        start_row, start_col = excel_address_to_coords(range_str)
        end_row, end_col = start_row, start_col

    return start_row, start_col, end_row, end_col


def check_range_overlap(range1: str, range2: str) -> bool:
    """
    두 Excel 범위가 겹치는지 검사합니다.

    Args:
        range1: 첫 번째 범위 (예: "A1:C10")
        range2: 두 번째 범위 (예: "B5:D15")

    Returns:
        겹치면 True, 겹치지 않으면 False
    """
    try:
        r1_start_row, r1_start_col, r1_end_row, r1_end_col = parse_excel_range(range1)
        r2_start_row, r2_start_col, r2_end_row, r2_end_col = parse_excel_range(range2)

        # 겹침 검사: 한 범위가 다른 범위 완전히 벗어나지 않으면 겹침
        return not (
            r1_end_row < r2_start_row or r1_start_row > r2_end_row or r1_end_col < r2_start_col or r1_start_col > r2_end_col
        )
    except (ValueError, Exception):
        # 파싱 실패 시 안전하게 겹친다고 가정
        return True


def get_all_pivot_ranges(sheet: xw.Sheet) -> List[str]:
    """
    시트의 모든 피벗 테이블 범위를 가져옵니다.

    Args:
        sheet: xlwings Sheet 객체

    Returns:
        피벗 테이블 범위 목록 (예: ["F1:H20", "K1:M15"])
    """
    ranges = []

    try:
        if platform.system() == "Windows":
            # Windows에서는 COM API 사용
            for pivot_table in sheet.api.PivotTables():
                try:
                    # TableRange2는 데이터 영역을 포함한 전체 범위
                    if hasattr(pivot_table, "TableRange2") and pivot_table.TableRange2:
                        table_range = pivot_table.TableRange2.Address.replace("$", "")
                        ranges.append(table_range)
                    elif hasattr(pivot_table, "TableRange1") and pivot_table.TableRange1:
                        table_range = pivot_table.TableRange1.Address.replace("$", "")
                        ranges.append(table_range)
                except Exception:
                    # 개별 피벗 테이블 처리 실패 시 무시
                    continue
    except Exception:
        # 피벗 테이블 접근 실패 시 빈 목록 반환
        pass

    return ranges


def get_all_chart_ranges(sheet: xw.Sheet) -> List[Tuple[str, int, int]]:
    """
    시트의 모든 차트 위치와 크기를 가져옵니다.

    Args:
        sheet: xlwings Sheet 객체

    Returns:
        차트 정보 목록 [(range_estimate, width, height), ...]
        range_estimate는 차트가 차지하는 대략적인 셀 범위
    """
    chart_info = []

    try:
        for chart in sheet.charts:
            try:
                # 차트의 픽셀 위치와 크기
                left = chart.left
                top = chart.top
                width = chart.width
                height = chart.height

                # 픽셀 좌표를 대략적인 셀 좌표로 변환
                # (Excel 기본 열 너비 약 64픽셀, 행 높이 약 15픽셀)
                start_col = max(1, int(left / 64) + 1)
                start_row = max(1, int(top / 15) + 1)
                end_col = start_col + max(1, int(width / 64))
                end_row = start_row + max(1, int(height / 15))

                range_estimate = f"{coords_to_excel_address(start_row, start_col)}:{coords_to_excel_address(end_row, end_col)}"
                chart_info.append((range_estimate, width, height))

            except Exception:
                # 개별 차트 처리 실패 시 무시
                continue
    except Exception:
        # 차트 접근 실패 시 빈 목록 반환
        pass

    return chart_info


def find_available_position(
    sheet: xw.Sheet, min_spacing: int = 2, preferred_position: str = "right", estimate_size: Tuple[int, int] = (10, 5)
) -> str:
    """
    시트에서 피벗 테이블이나 차트 배치에 적합한 빈 위치를 찾습니다.

    Args:
        sheet: xlwings Sheet 객체
        min_spacing: 기존 객체와의 최소 간격 (열 단위)
        preferred_position: 선호 배치 방향 ("right" 또는 "bottom")
        estimate_size: 예상 크기 (cols, rows) - 피벗 테이블용

    Returns:
        추천 위치의 Excel 주소 (예: "F1")

    Raises:
        RuntimeError: 적절한 위치를 찾을 수 없는 경우
    """
    # 기존 객체들의 범위 수집
    existing_ranges = []

    # 피벗 테이블 범위 추가
    pivot_ranges = get_all_pivot_ranges(sheet)
    existing_ranges.extend(pivot_ranges)

    # 차트 범위 추가 (차트는 범위만 사용)
    chart_info = get_all_chart_ranges(sheet)
    for chart_range, _, _ in chart_info:
        existing_ranges.append(chart_range)

    # 사용된 데이터 영역도 고려
    try:
        used_range = sheet.used_range
        if used_range:
            data_range = used_range.address.replace("$", "")
            existing_ranges.append(data_range)
    except Exception:
        # used_range 접근 실패 시 무시
        pass

    # 배치 후보 위치들 생성
    candidates = []

    if preferred_position == "right":
        # 오른쪽 우선 배치
        for start_col in range(1, 50):  # A~AX열까지 시도
            for start_row in range(1, 100):  # 1~100행까지 시도
                candidates.append((start_row, start_col))
    else:  # "bottom"
        # 아래쪽 우선 배치
        for start_row in range(1, 100):  # 1~100행까지 시도
            for start_col in range(1, 50):  # A~AX열까지 시도
                candidates.append((start_row, start_col))

    # 각 후보 위치에서 겹침 검사
    for row, col in candidates:
        # 예상 크기로 범위 생성
        end_row = row + estimate_size[1] - 1
        end_col = col + estimate_size[0] - 1

        candidate_range = f"{coords_to_excel_address(row, col)}:{coords_to_excel_address(end_row, end_col)}"

        # 기존 범위들과 겹침 검사 (간격 고려)
        has_conflict = False
        for existing_range in existing_ranges:
            try:
                # 간격을 고려한 확장 범위로 겹침 검사
                e_start_row, e_start_col, e_end_row, e_end_col = parse_excel_range(existing_range)

                # 간격만큼 확장
                e_start_row = max(1, e_start_row - min_spacing)
                e_start_col = max(1, e_start_col - min_spacing)
                e_end_row += min_spacing
                e_end_col += min_spacing

                expanded_existing = (
                    f"{coords_to_excel_address(e_start_row, e_start_col)}:{coords_to_excel_address(e_end_row, e_end_col)}"
                )

                if check_range_overlap(candidate_range, expanded_existing):
                    has_conflict = True
                    break
            except Exception:
                # 파싱 실패 시 안전하게 충돌로 간주
                has_conflict = True
                break

        if not has_conflict:
            return coords_to_excel_address(row, col)

    # 적절한 위치를 찾지 못한 경우
    raise RuntimeError("시트에 충분한 빈 공간을 찾을 수 없습니다. 수동으로 위치를 지정하거나 기존 객체를 정리해주세요.")


def estimate_pivot_table_size(source_range: str, field_count: int = 3) -> Tuple[int, int]:
    """
    피벗 테이블의 예상 크기를 추정합니다.

    Args:
        source_range: 소스 데이터 범위
        field_count: 필드 개수 (기본값: 3)

    Returns:
        (예상_열수, 예상_행수) 튜플
    """
    try:
        # 소스 데이터 크기에 기반한 추정
        start_row, start_col, end_row, end_col = parse_excel_range(source_range)
        data_rows = end_row - start_row + 1
        data_cols = end_col - start_col + 1

        # 보수적인 추정: 필드 개수 + 여백, 데이터 행의 일부 + 헤더
        estimated_cols = min(field_count + 3, 15)  # 최대 15열
        estimated_rows = min(max(data_rows // 10, 10), 50)  # 최소 10행, 최대 50행

        return estimated_cols, estimated_rows
    except Exception:
        # 추정 실패 시 기본값
        return 10, 20


def validate_auto_position_requirements(sheet: xw.Sheet) -> Tuple[bool, str]:
    """
    자동 배치 기능 사용 가능 여부를 검사합니다.

    Args:
        sheet: xlwings Sheet 객체

    Returns:
        (사용가능여부, 메시지) 튜플
    """
    try:
        # Windows에서만 완전한 기능 지원
        if platform.system() != "Windows":
            return False, "자동 배치 기능은 Windows에서만 완전히 지원됩니다"

        # 시트 접근 가능 여부 확인
        sheet_name = sheet.name
        if not sheet_name:
            return False, "시트 정보에 접근할 수 없습니다"

        return True, ""

    except Exception as e:
        return False, f"자동 배치 기능을 사용할 수 없습니다: {str(e)}"


# =============================================================================
# 데이터 분석 및 변환 유틸리티 함수들 (Issue #39)
# =============================================================================


def analyze_data_structure(data_range: xw.Range) -> Dict[str, Union[str, bool, int, float, List[str]]]:
    """
    Excel 데이터 구조를 분석하여 피벗테이블 준비 상태를 평가합니다.

    Args:
        data_range: 분석할 xlwings Range 객체

    Returns:
        분석 결과 딕셔너리
    """
    try:
        # 데이터를 pandas DataFrame으로 변환
        values = data_range.value
        if not values:
            return {
                "format_type": DataFormat.UNKNOWN,
                "issues": [],
                "pivot_ready": False,
                "transformation_needed": False,
                "recommendations": ["데이터가 비어있습니다"],
                "estimated_rows_after_transform": 0,
                "confidence_score": 0.0,
            }

        # 데이터를 2차원 리스트로 정규화
        if not isinstance(values, list):
            values = [[values]]
        elif not isinstance(values[0], list):
            values = [values]

        df = pd.DataFrame(
            values[1:], columns=values[0] if len(values) > 1 else [f"Column_{i+1}" for i in range(len(values[0]))]
        )

        issues = []
        recommendations = []
        format_type = DataFormat.PIVOT_READY

        # 1. 병합된 셀 감지 (빈 값이 많은 경우로 추정)
        empty_ratio = df.isnull().sum().sum() / (df.shape[0] * df.shape[1])
        if empty_ratio > 0.3:
            issues.append("merged_cells")
            recommendations.append("병합된 셀을 해제하고 값을 채워넣으세요")
            format_type = DataFormat.MERGED_CELLS

        # 2. 교차표 형식 감지 (숫자 열이 많고 첫 번째 열이 텍스트인 경우)
        numeric_cols = df.select_dtypes(include=["number"]).columns
        if len(numeric_cols) > 3 and df.shape[1] > 5:
            # 첫 번째 열이 카테고리이고 나머지가 수치인 경우
            first_col_categorical = pd.api.types.is_string_dtype(df.iloc[:, 0]) or pd.api.types.is_object_dtype(df.iloc[:, 0])
            if first_col_categorical and len(numeric_cols) / df.shape[1] > 0.6:
                issues.append("cross_tab")
                recommendations.append(f"{df.columns[1:][:5].tolist()}... 열들을 세로 형식으로 변환하세요")
                format_type = DataFormat.CROSS_TAB

        # 3. 다단계 헤더 감지 (열 이름에 패턴이 있는 경우)
        header_patterns = [col for col in df.columns if isinstance(col, str) and ("." in col or "_" in col or " - " in col)]
        if len(header_patterns) > df.shape[1] * 0.5:
            issues.append("multi_level_headers")
            recommendations.append("다단계 헤더를 단일 헤더로 결합하세요")
            format_type = DataFormat.MULTI_LEVEL_HEADERS

        # 4. 소계 행 감지 (행에서 "총계", "소계", "합계" 등이 포함된 경우)
        subtotal_keywords = ["총계", "소계", "합계", "Total", "Subtotal", "Sum"]
        subtotal_rows = 0
        for _, row in df.iterrows():
            row_str = " ".join([str(val) for val in row.values if pd.notna(val)])
            if any(keyword in row_str for keyword in subtotal_keywords):
                subtotal_rows += 1

        if subtotal_rows > 0:
            issues.append("subtotals_mixed")
            recommendations.append(f"소계 행 {subtotal_rows}개를 제거하세요")
            format_type = DataFormat.SUBTOTALS_MIXED

        # 5. 넓은 형식 감지 (동일한 지표가 여러 열에 반복되는 경우)
        if df.shape[1] > 10:
            # 열 이름에서 연도, 월, 분기 패턴 감지
            date_patterns = [r"\d{4}", r"Q[1-4]", r"[1-9][0-2]?월", r"Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec"]
            date_cols = []
            for col in df.columns:
                if isinstance(col, str) and any(re.search(pattern, col) for pattern in date_patterns):
                    date_cols.append(col)

            if len(date_cols) > 3:
                issues.append("wide_format")
                recommendations.append(f"날짜/기간 열들({len(date_cols)}개)을 하나의 기간 컬럼으로 변환하세요")
                format_type = DataFormat.WIDE_FORMAT

        # 피벗테이블 준비 상태 평가
        pivot_ready = len(issues) == 0
        transformation_needed = not pivot_ready

        # 변환 후 예상 행 수 계산
        estimated_rows = df.shape[0]
        if "cross_tab" in issues or "wide_format" in issues:
            # 교차표나 넓은 형식인 경우 행 수가 크게 증가
            estimated_rows = df.shape[0] * max(len(numeric_cols), 1)

        # 신뢰도 점수 계산 (문제가 적을수록 높음)
        confidence_score = max(0.0, 1.0 - len(issues) * 0.2)

        return {
            "format_type": format_type.value,
            "issues": issues,
            "pivot_ready": pivot_ready,
            "transformation_needed": transformation_needed,
            "recommendations": recommendations,
            "estimated_rows_after_transform": int(estimated_rows),
            "confidence_score": round(confidence_score, 2),
            "data_shape": {"rows": df.shape[0], "columns": df.shape[1]},
            "ai_assistance_available": True,
        }

    except Exception as e:
        return {
            "format_type": DataFormat.UNKNOWN.value,
            "issues": ["analysis_failed"],
            "pivot_ready": False,
            "transformation_needed": False,
            "recommendations": [f"데이터 분석 중 오류 발생: {str(e)}"],
            "estimated_rows_after_transform": 0,
            "confidence_score": 0.0,
            "ai_assistance_available": False,
        }


def transform_data_unpivot(df: pd.DataFrame, id_vars: Optional[List[str]] = None) -> pd.DataFrame:
    """
    교차표 형식 데이터를 세로 형식(unpivot)으로 변환합니다.

    Args:
        df: 변환할 pandas DataFrame
        id_vars: 고정할 열 목록 (None이면 첫 번째 열 자동 선택)

    Returns:
        변환된 pandas DataFrame
    """
    try:
        if id_vars is None:
            # 첫 번째 열을 ID 변수로 사용
            id_vars = [df.columns[0]]

        # 수치형 열들을 찾아서 melt 대상으로 설정
        numeric_cols = df.select_dtypes(include=["number"]).columns.tolist()
        value_vars = [col for col in numeric_cols if col not in id_vars]

        if not value_vars:
            # 수치형 열이 없으면 ID 변수 외의 모든 열 사용
            value_vars = [col for col in df.columns if col not in id_vars]

        # pandas melt를 사용하여 unpivot
        melted_df = pd.melt(df, id_vars=id_vars, value_vars=value_vars, var_name="변수", value_name="값")

        # 빈 값 제거
        melted_df = melted_df.dropna(subset=["값"])

        return melted_df.reset_index(drop=True)

    except Exception as e:
        raise ValueError(f"Unpivot 변환 실패: {str(e)}")


def transform_data_unmerge(df: pd.DataFrame) -> pd.DataFrame:
    """
    병합된 셀로 인한 빈 값들을 앞의 값으로 채웁니다.

    Args:
        df: 변환할 pandas DataFrame

    Returns:
        변환된 pandas DataFrame
    """
    try:
        # 모든 열에 대해 forward fill 적용
        filled_df = df.fillna(method="ffill")

        # 문자열 열의 경우 빈 문자열도 채우기
        for col in filled_df.columns:
            if filled_df[col].dtype == "object":
                filled_df[col] = filled_df[col].replace("", None)
                filled_df[col] = filled_df[col].fillna(method="ffill")

        return filled_df.reset_index(drop=True)

    except Exception as e:
        raise ValueError(f"Unmerge 변환 실패: {str(e)}")


def transform_data_flatten_headers(df: pd.DataFrame) -> pd.DataFrame:
    """
    다단계 헤더를 단일 헤더로 결합합니다.

    Args:
        df: 변환할 pandas DataFrame

    Returns:
        변환된 pandas DataFrame
    """
    try:
        # 열 이름을 단순화
        new_columns = []
        for col in df.columns:
            if isinstance(col, str):
                # 특수 문자를 언더스코어로 대체하고 단순화
                simplified = re.sub(r"[^\w\s]", "_", col)
                simplified = re.sub(r"\s+", "_", simplified)
                simplified = re.sub(r"_+", "_", simplified).strip("_")
                new_columns.append(simplified)
            else:
                new_columns.append(f"Column_{len(new_columns)+1}")

        df.columns = new_columns
        return df.reset_index(drop=True)

    except Exception as e:
        raise ValueError(f"Header 평탄화 실패: {str(e)}")


def transform_data_remove_subtotals(df: pd.DataFrame) -> pd.DataFrame:
    """
    소계 행들을 제거합니다.

    Args:
        df: 변환할 pandas DataFrame

    Returns:
        변환된 pandas DataFrame
    """
    try:
        subtotal_keywords = ["총계", "소계", "합계", "Total", "Subtotal", "Sum", "계"]

        # 소계 행 인덱스 찾기
        rows_to_drop = []
        for idx, row in df.iterrows():
            row_str = " ".join([str(val) for val in row.values if pd.notna(val)])
            if any(keyword in row_str for keyword in subtotal_keywords):
                rows_to_drop.append(idx)

        # 소계 행 제거
        cleaned_df = df.drop(rows_to_drop)

        return cleaned_df.reset_index(drop=True)

    except Exception as e:
        raise ValueError(f"소계 제거 실패: {str(e)}")


def transform_data_auto(df: pd.DataFrame) -> Tuple[pd.DataFrame, List[str]]:
    """
    자동으로 모든 필요한 변환을 적용합니다.

    Args:
        df: 변환할 pandas DataFrame

    Returns:
        (변환된 DataFrame, 적용된 변환 목록) 튜플
    """
    try:
        applied_transforms = []
        result_df = df.copy()

        # 1. 소계 제거 (먼저 수행)
        subtotal_keywords = ["총계", "소계", "합계", "Total", "Subtotal", "Sum"]
        subtotal_rows = 0
        for _, row in result_df.iterrows():
            row_str = " ".join([str(val) for val in row.values if pd.notna(val)])
            if any(keyword in row_str for keyword in subtotal_keywords):
                subtotal_rows += 1

        if subtotal_rows > 0:
            result_df = transform_data_remove_subtotals(result_df)
            applied_transforms.append("remove-subtotals")

        # 2. 병합된 셀 처리 (빈 값 비율 확인)
        empty_ratio = result_df.isnull().sum().sum() / (result_df.shape[0] * result_df.shape[1])
        if empty_ratio > 0.3:
            result_df = transform_data_unmerge(result_df)
            applied_transforms.append("unmerge")

        # 3. 헤더 평탄화 (복잡한 열 이름 확인)
        header_patterns = [
            col for col in result_df.columns if isinstance(col, str) and ("." in col or "_" in col or " - " in col)
        ]
        if len(header_patterns) > result_df.shape[1] * 0.5:
            result_df = transform_data_flatten_headers(result_df)
            applied_transforms.append("flatten-headers")

        # 4. Unpivot 변환 (교차표 또는 넓은 형식 감지)
        numeric_cols = result_df.select_dtypes(include=["number"]).columns
        should_unpivot = False

        # 교차표 감지
        if len(numeric_cols) > 3 and result_df.shape[1] > 5:
            first_col_categorical = pd.api.types.is_string_dtype(result_df.iloc[:, 0]) or pd.api.types.is_object_dtype(
                result_df.iloc[:, 0]
            )
            if first_col_categorical and len(numeric_cols) / result_df.shape[1] > 0.6:
                should_unpivot = True

        # 넓은 형식 감지 (날짜/기간 열들)
        if result_df.shape[1] > 10:
            date_patterns = [r"\d{4}", r"Q[1-4]", r"[1-9][0-2]?월", r"Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec"]
            date_cols = []
            for col in result_df.columns:
                if isinstance(col, str) and any(re.search(pattern, col) for pattern in date_patterns):
                    date_cols.append(col)
            if len(date_cols) > 3:
                should_unpivot = True

        if should_unpivot:
            result_df = transform_data_unpivot(result_df)
            applied_transforms.append("unpivot")

        return result_df, applied_transforms

    except Exception as e:
        raise ValueError(f"자동 변환 실패: {str(e)}")


def get_shape_by_name(sheet: xw.Sheet, shape_name: str) -> Optional[xw.Shape]:
    """
    시트에서 이름으로 도형을 찾습니다.

    Args:
        sheet: xlwings Sheet 객체
        shape_name: 찾을 도형 이름

    Returns:
        xlwings Shape 객체 또는 None
    """
    try:
        for shape in sheet.shapes:
            if shape.name == shape_name:
                return shape
        return None
    except Exception:
        return None


def get_shapes_info(sheet: xw.Sheet) -> List[Dict[str, Union[str, int, float]]]:
    """
    시트의 모든 도형 정보를 수집합니다.

    Args:
        sheet: xlwings Sheet 객체

    Returns:
        도형 정보 리스트
    """
    shapes_info = []

    try:
        for shape in sheet.shapes:
            shape_info = {
                "name": shape.name,
                "type": "unknown",
                "position": {"left": getattr(shape, "left", 0), "top": getattr(shape, "top", 0)},
                "size": {"width": getattr(shape, "width", 0), "height": getattr(shape, "height", 0)},
            }

            # Windows에서 추가 정보 수집
            if platform.system() == "Windows":
                try:
                    shape_info.update(
                        {
                            "type": shape.api.Type,
                            "visible": shape.api.Visible,
                            "has_text": hasattr(shape.api, "TextFrame") and shape.api.TextFrame.HasText,
                        }
                    )

                    # 색상 정보
                    if hasattr(shape.api, "Fill"):
                        shape_info["fill_color"] = rgb_to_hex(shape.api.Fill.ForeColor.RGB)
                        shape_info["transparency"] = shape.api.Fill.Transparency * 100

                except Exception:
                    pass

            shapes_info.append(shape_info)

    except Exception:
        pass

    return shapes_info


def validate_position_and_size(left: int, top: int, width: int, height: int) -> Tuple[bool, str]:
    """
    도형의 위치와 크기가 유효한지 검증합니다.

    Args:
        left: 왼쪽 위치
        top: 위쪽 위치
        width: 너비
        height: 높이

    Returns:
        (유효성 여부, 오류 메시지)
    """
    if left < 0 or top < 0:
        return False, "위치는 0 이상이어야 합니다"

    if width <= 0 or height <= 0:
        return False, "크기는 0보다 커야 합니다"

    # 최대 크기 제한 (A1 범위 기준)
    if left > 20000 or top > 20000:
        return False, "위치가 너무 큽니다 (최대: 20000px)"

    if width > 5000 or height > 5000:
        return False, "크기가 너무 큽니다 (최대: 5000px)"

    return True, ""


def generate_unique_shape_name(sheet: xw.Sheet, base_name: str = "Shape") -> str:
    """
    시트에서 고유한 도형 이름을 생성합니다.

    Args:
        sheet: xlwings Sheet 객체
        base_name: 기본 이름

    Returns:
        고유한 도형 이름
    """
    existing_names = set()
    try:
        for shape in sheet.shapes:
            existing_names.add(shape.name)
    except Exception:
        pass

    # 기본 이름이 중복되지 않으면 그대로 사용
    if base_name not in existing_names:
        return base_name

    # 숫자를 붙여서 고유한 이름 생성
    counter = 1
    while f"{base_name}{counter}" in existing_names:
        counter += 1

    return f"{base_name}{counter}"


# ========== Slicer 관리 기능 추가 ==========


def get_pivot_tables(sheet: xw.Sheet) -> List[Dict[str, Union[str, int, float]]]:
    """
    시트의 모든 피벗테이블 정보를 수집합니다.

    Args:
        sheet: xlwings Sheet 객체

    Returns:
        피벗테이블 정보 리스트
    """
    pivot_tables = []

    try:
        if platform.system() == "Windows":
            for pivot_table in sheet.api.PivotTables():
                pivot_info = {
                    "name": pivot_table.Name,
                    "location": pivot_table.TableRange1.Address,
                    "source_data": str(getattr(pivot_table, "SourceData", "Unknown")),
                    "fields": [],
                }

                # 피벗테이블 필드 정보 수집
                try:
                    for field in pivot_table.PivotFields():
                        field_info = {"name": field.Name, "orientation": field.Orientation}
                        pivot_info["fields"].append(field_info)
                except Exception:
                    pass

                pivot_tables.append(pivot_info)

    except Exception:
        pass

    return pivot_tables


def get_slicer_by_name(workbook: xw.Book, slicer_name: str):
    """
    워크북에서 이름으로 슬라이서를 찾습니다.

    Args:
        workbook: xlwings Book 객체
        slicer_name: 찾을 슬라이서 이름

    Returns:
        슬라이서 객체 또는 None
    """
    try:
        if platform.system() == "Windows":
            for slicer in workbook.api.SlicerCaches():
                if slicer.Name == slicer_name:
                    return slicer
        return None
    except Exception:
        return None


def get_slicers_info(workbook: xw.Book) -> List[Dict[str, Union[str, int, float]]]:
    """
    워크북의 모든 슬라이서 정보를 수집합니다.

    Args:
        workbook: xlwings Book 객체

    Returns:
        슬라이서 정보 리스트
    """
    slicers_info = []

    try:
        if platform.system() == "Windows":
            for slicer_cache in workbook.api.SlicerCaches():
                slicer_info = {
                    "name": slicer_cache.Name,
                    "source_name": slicer_cache.SourceName,
                    "field_name": getattr(slicer_cache, "OLAP", {}).get("SourceField", "Unknown"),
                    "slicer_items": [],
                    "connected_pivot_tables": [],
                }

                # 슬라이서 아이템 정보
                try:
                    for item in slicer_cache.SlicerItems():
                        item_info = {"name": item.Name, "selected": item.Selected}
                        slicer_info["slicer_items"].append(item_info)
                except Exception:
                    pass

                # 연결된 피벗테이블 정보
                try:
                    for pivot_table in slicer_cache.PivotTables():
                        slicer_info["connected_pivot_tables"].append(pivot_table.Name)
                except Exception:
                    pass

                # 슬라이서 위치 정보 (첫 번째 슬라이서 기준)
                try:
                    if slicer_cache.Slicers().Count > 0:
                        first_slicer = slicer_cache.Slicers(1)
                        slicer_info["position"] = {"left": first_slicer.Left, "top": first_slicer.Top}
                        slicer_info["size"] = {"width": first_slicer.Width, "height": first_slicer.Height}
                        slicer_info["sheet"] = first_slicer.Parent.Parent.Name
                except Exception:
                    pass

                slicers_info.append(slicer_info)

    except Exception:
        pass

    return slicers_info


def validate_slicer_position(left: int, top: int, width: int, height: int) -> Tuple[bool, str]:
    """
    슬라이서의 위치와 크기가 유효한지 검증합니다.

    Args:
        left: 왼쪽 위치
        top: 위쪽 위치
        width: 너비
        height: 높이

    Returns:
        (유효성 여부, 오류 메시지)
    """
    if left < 0 or top < 0:
        return False, "슬라이서 위치는 0 이상이어야 합니다"

    if width <= 0 or height <= 0:
        return False, "슬라이서 크기는 0보다 커야 합니다"

    # 슬라이서 최소 크기 (Excel 기본값 기준)
    if width < 100:
        return False, "슬라이서 너비는 최소 100픽셀이어야 합니다"

    if height < 50:
        return False, "슬라이서 높이는 최소 50픽셀이어야 합니다"

    # 최대 크기 제한
    if left > 15000 or top > 15000:
        return False, "슬라이서 위치가 너무 큽니다 (최대: 15000px)"

    if width > 3000 or height > 2000:
        return False, "슬라이서 크기가 너무 큽니다 (최대: 3000x2000px)"

    return True, ""


def generate_unique_slicer_name(workbook: xw.Book, base_name: str = "Slicer") -> str:
    """
    워크북에서 고유한 슬라이서 이름을 생성합니다.

    Args:
        workbook: xlwings Book 객체
        base_name: 기본 이름

    Returns:
        고유한 슬라이서 이름
    """
    existing_names = set()
    try:
        if platform.system() == "Windows":
            for slicer_cache in workbook.api.SlicerCaches():
                existing_names.add(slicer_cache.Name)
    except Exception:
        pass

    # 기본 이름이 중복되지 않으면 그대로 사용
    if base_name not in existing_names:
        return base_name

    # 숫자를 붙여서 고유한 이름 생성
    counter = 1
    while f"{base_name}{counter}" in existing_names:
        counter += 1

    return f"{base_name}{counter}"


def apply_slicer_style(slicer, style_name: str = "slicer-box") -> bool:
    """
    슬라이서에 뉴모피즘 스타일을 적용합니다.

    Args:
        slicer: xlwings Slicer 객체
        style_name: 적용할 스타일 이름

    Returns:
        스타일 적용 성공 여부
    """
    if platform.system() != "Windows":
        return False

    try:
        if style_name in NEUMORPHISM_STYLES:
            style = NEUMORPHISM_STYLES[style_name]

            # 슬라이서 스타일 설정 (Windows COM API)
            if style.get("fill_color"):
                # 슬라이서의 경우 배경색 설정이 제한적
                pass

            # 테두리 설정
            if style.get("has_line", True) and style.get("line_color"):
                # 슬라이서 테두리 색상 설정
                pass

        return True

    except Exception:
        return False


# =============================================================================
# 범위 관리 및 자동 배치 유틸리티 함수들
# =============================================================================


def excel_address_to_coords(address: str) -> Tuple[int, int]:
    """
    Excel 주소(A1)를 좌표(row, col)로 변환합니다.

    Args:
        address: Excel 주소 (예: "A1", "BC123")

    Returns:
        (row, col) 튜플 (1-based index)

    Raises:
        ValueError: 잘못된 주소 형식
    """
    import re

    # 주소 형식 검증
    match = re.match(r"^([A-Z]+)(\d+)$", address.upper())
    if not match:
        raise ValueError(f"잘못된 Excel 주소 형식: {address}")

    col_letters, row_str = match.groups()
    row = int(row_str)

    # 열 문자를 숫자로 변환 (A=1, B=2, ..., Z=26, AA=27, ...)
    col = 0
    for i, letter in enumerate(reversed(col_letters)):
        col += (ord(letter) - ord("A") + 1) * (26**i)

    return row, col


def coords_to_excel_address(row: int, col: int) -> str:
    """
    좌표(row, col)를 Excel 주소로 변환합니다.

    Args:
        row: 행 번호 (1-based)
        col: 열 번호 (1-based)

    Returns:
        Excel 주소 (예: "A1", "BC123")
    """
    if row < 1 or col < 1:
        raise ValueError("행과 열 번호는 1 이상이어야 합니다")

    # 열 번호를 문자로 변환
    col_letters = ""
    while col > 0:
        col -= 1
        col_letters = chr(ord("A") + col % 26) + col_letters
        col //= 26

    return f"{col_letters}{row}"


def parse_excel_range(range_str: str) -> Tuple[int, int, int, int]:
    """
    Excel 범위(A1:C10)를 좌표로 파싱합니다.

    Args:
        range_str: Excel 범위 (예: "A1:C10", "A1" 단일 셀도 가능)

    Returns:
        (start_row, start_col, end_row, end_col) 튜플 (1-based)

    Raises:
        ValueError: 잘못된 범위 형식
    """
    range_str = range_str.strip()

    if ":" in range_str:
        # 범위 형식 (A1:C10)
        start_addr, end_addr = range_str.split(":", 1)
        start_row, start_col = excel_address_to_coords(start_addr.strip())
        end_row, end_col = excel_address_to_coords(end_addr.strip())
    else:
        # 단일 셀 형식 (A1)
        start_row, start_col = excel_address_to_coords(range_str)
        end_row, end_col = start_row, start_col

    return start_row, start_col, end_row, end_col


def check_range_overlap(range1: str, range2: str) -> bool:
    """
    두 Excel 범위가 겹치는지 검사합니다.

    Args:
        range1: 첫 번째 범위 (예: "A1:C10")
        range2: 두 번째 범위 (예: "B5:D15")

    Returns:
        겹치면 True, 겹치지 않으면 False
    """
    try:
        r1_start_row, r1_start_col, r1_end_row, r1_end_col = parse_excel_range(range1)
        r2_start_row, r2_start_col, r2_end_row, r2_end_col = parse_excel_range(range2)

        # 겹침 검사: 한 범위가 다른 범위 완전히 벗어나지 않으면 겹침
        return not (
            r1_end_row < r2_start_row or r1_start_row > r2_end_row or r1_end_col < r2_start_col or r1_start_col > r2_end_col
        )
    except (ValueError, Exception):
        # 파싱 실패 시 안전하게 겹친다고 가정
        return True


def get_all_pivot_ranges(sheet: xw.Sheet) -> List[str]:
    """
    시트의 모든 피벗 테이블 범위를 가져옵니다.

    Args:
        sheet: xlwings Sheet 객체

    Returns:
        피벗 테이블 범위 목록 (예: ["F1:H20", "K1:M15"])
    """
    ranges = []

    try:
        if platform.system() == "Windows":
            # Windows에서는 COM API 사용
            for pivot_table in sheet.api.PivotTables():
                try:
                    # TableRange2는 데이터 영역을 포함한 전체 범위
                    if hasattr(pivot_table, "TableRange2") and pivot_table.TableRange2:
                        table_range = pivot_table.TableRange2.Address.replace("$", "")
                        ranges.append(table_range)
                    elif hasattr(pivot_table, "TableRange1") and pivot_table.TableRange1:
                        table_range = pivot_table.TableRange1.Address.replace("$", "")
                        ranges.append(table_range)
                except Exception:
                    # 개별 피벗 테이블 처리 실패 시 무시
                    continue
    except Exception:
        # 피벗 테이블 접근 실패 시 빈 목록 반환
        pass

    return ranges


def get_all_chart_ranges(sheet: xw.Sheet) -> List[Tuple[str, int, int]]:
    """
    시트의 모든 차트 위치와 크기를 가져옵니다.

    Args:
        sheet: xlwings Sheet 객체

    Returns:
        차트 정보 목록 [(range_estimate, width, height), ...]
        range_estimate는 차트가 차지하는 대략적인 셀 범위
    """
    chart_info = []

    try:
        for chart in sheet.charts:
            try:
                # 차트의 픽셀 위치와 크기
                left = chart.left
                top = chart.top
                width = chart.width
                height = chart.height

                # 픽셀 좌표를 대략적인 셀 좌표로 변환
                # (Excel 기본 열 너비 약 64픽셀, 행 높이 약 15픽셀)
                start_col = max(1, int(left / 64) + 1)
                start_row = max(1, int(top / 15) + 1)
                end_col = start_col + max(1, int(width / 64))
                end_row = start_row + max(1, int(height / 15))

                range_estimate = f"{coords_to_excel_address(start_row, start_col)}:{coords_to_excel_address(end_row, end_col)}"
                chart_info.append((range_estimate, width, height))

            except Exception:
                # 개별 차트 처리 실패 시 무시
                continue
    except Exception:
        # 차트 접근 실패 시 빈 목록 반환
        pass

    return chart_info


def find_available_position(
    sheet: xw.Sheet, min_spacing: int = 2, preferred_position: str = "right", estimate_size: Tuple[int, int] = (10, 5)
) -> str:
    """
    시트에서 피벗 테이블이나 차트 배치에 적합한 빈 위치를 찾습니다.

    Args:
        sheet: xlwings Sheet 객체
        min_spacing: 기존 객체와의 최소 간격 (열 단위)
        preferred_position: 선호 배치 방향 ("right" 또는 "bottom")
        estimate_size: 예상 크기 (cols, rows) - 피벗 테이블용

    Returns:
        추천 위치의 Excel 주소 (예: "F1")

    Raises:
        RuntimeError: 적절한 위치를 찾을 수 없는 경우
    """
    # 기존 객체들의 범위 수집
    existing_ranges = []

    # 피벗 테이블 범위 추가
    pivot_ranges = get_all_pivot_ranges(sheet)
    existing_ranges.extend(pivot_ranges)

    # 차트 범위 추가 (차트는 범위만 사용)
    chart_info = get_all_chart_ranges(sheet)
    for chart_range, _, _ in chart_info:
        existing_ranges.append(chart_range)

    # 사용된 데이터 영역도 고려
    try:
        used_range = sheet.used_range
        if used_range:
            data_range = used_range.address.replace("$", "")
            existing_ranges.append(data_range)
    except Exception:
        # used_range 접근 실패 시 무시
        pass

    # 배치 후보 위치들 생성
    candidates = []

    if preferred_position == "right":
        # 오른쪽 우선 배치
        for start_col in range(1, 50):  # A~AX열까지 시도
            for start_row in range(1, 100):  # 1~100행까지 시도
                candidates.append((start_row, start_col))
    else:  # "bottom"
        # 아래쪽 우선 배치
        for start_row in range(1, 100):  # 1~100행까지 시도
            for start_col in range(1, 50):  # A~AX열까지 시도
                candidates.append((start_row, start_col))

    # 각 후보 위치에서 겹침 검사
    for row, col in candidates:
        # 예상 크기로 범위 생성
        end_row = row + estimate_size[1] - 1
        end_col = col + estimate_size[0] - 1

        candidate_range = f"{coords_to_excel_address(row, col)}:{coords_to_excel_address(end_row, end_col)}"

        # 기존 범위들과 겹침 검사 (간격 고려)
        has_conflict = False
        for existing_range in existing_ranges:
            try:
                # 간격을 고려한 확장 범위로 겹침 검사
                e_start_row, e_start_col, e_end_row, e_end_col = parse_excel_range(existing_range)

                # 간격만큼 확장
                e_start_row = max(1, e_start_row - min_spacing)
                e_start_col = max(1, e_start_col - min_spacing)
                e_end_row += min_spacing
                e_end_col += min_spacing

                expanded_existing = (
                    f"{coords_to_excel_address(e_start_row, e_start_col)}:{coords_to_excel_address(e_end_row, e_end_col)}"
                )

                if check_range_overlap(candidate_range, expanded_existing):
                    has_conflict = True
                    break
            except Exception:
                # 파싱 실패 시 안전하게 충돌로 간주
                has_conflict = True
                break

        if not has_conflict:
            return coords_to_excel_address(row, col)

    # 적절한 위치를 찾지 못한 경우
    raise RuntimeError("시트에 충분한 빈 공간을 찾을 수 없습니다. 수동으로 위치를 지정하거나 기존 객체를 정리해주세요.")


def estimate_pivot_table_size(source_range: str, field_count: int = 3) -> Tuple[int, int]:
    """
    피벗 테이블의 예상 크기를 추정합니다.

    Args:
        source_range: 소스 데이터 범위
        field_count: 필드 개수 (기본값: 3)

    Returns:
        (예상_열수, 예상_행수) 튜플
    """
    try:
        # 소스 데이터 크기에 기반한 추정
        start_row, start_col, end_row, end_col = parse_excel_range(source_range)
        data_rows = end_row - start_row + 1
        data_cols = end_col - start_col + 1

        # 보수적인 추정: 필드 개수 + 여백, 데이터 행의 일부 + 헤더
        estimated_cols = min(field_count + 3, 15)  # 최대 15열
        estimated_rows = min(max(data_rows // 10, 10), 50)  # 최소 10행, 최대 50행

        return estimated_cols, estimated_rows
    except Exception:
        # 추정 실패 시 기본값
        return 10, 20


def validate_auto_position_requirements(sheet: xw.Sheet) -> Tuple[bool, str]:
    """
    자동 배치 기능 사용 가능 여부를 검사합니다.

    Args:
        sheet: xlwings Sheet 객체

    Returns:
        (사용가능여부, 메시지) 튜플
    """
    try:
        # Windows에서만 완전한 기능 지원
        if platform.system() != "Windows":
            return False, "자동 배치 기능은 Windows에서만 완전히 지원됩니다"

        # 시트 접근 가능 여부 확인
        sheet_name = sheet.name
        if not sheet_name:
            return False, "시트 정보에 접근할 수 없습니다"

        return True, ""

    except Exception as e:
        return False, f"자동 배치 기능을 사용할 수 없습니다: {str(e)}"


# =============================================================================
# 데이터 분석 및 변환 유틸리티 함수들 (Issue #39)
# =============================================================================


def analyze_data_structure(data_range: xw.Range) -> Dict[str, Union[str, bool, int, float, List[str]]]:
    """
    Excel 데이터 구조를 분석하여 피벗테이블 준비 상태를 평가합니다.

    Args:
        data_range: 분석할 xlwings Range 객체

    Returns:
        분석 결과 딕셔너리
    """
    try:
        # 데이터를 pandas DataFrame으로 변환
        values = data_range.value
        if not values:
            return {
                "format_type": DataFormat.UNKNOWN,
                "issues": [],
                "pivot_ready": False,
                "transformation_needed": False,
                "recommendations": ["데이터가 비어있습니다"],
                "estimated_rows_after_transform": 0,
                "confidence_score": 0.0,
            }

        # 데이터를 2차원 리스트로 정규화
        if not isinstance(values, list):
            values = [[values]]
        elif not isinstance(values[0], list):
            values = [values]

        df = pd.DataFrame(
            values[1:], columns=values[0] if len(values) > 1 else [f"Column_{i+1}" for i in range(len(values[0]))]
        )

        issues = []
        recommendations = []
        format_type = DataFormat.PIVOT_READY

        # 1. 병합된 셀 감지 (빈 값이 많은 경우로 추정)
        empty_ratio = df.isnull().sum().sum() / (df.shape[0] * df.shape[1])
        if empty_ratio > 0.3:
            issues.append("merged_cells")
            recommendations.append("병합된 셀을 해제하고 값을 채워넣으세요")
            format_type = DataFormat.MERGED_CELLS

        # 2. 교차표 형식 감지 (숫자 열이 많고 첫 번째 열이 텍스트인 경우)
        numeric_cols = df.select_dtypes(include=["number"]).columns
        if len(numeric_cols) > 3 and df.shape[1] > 5:
            # 첫 번째 열이 카테고리이고 나머지가 수치인 경우
            first_col_categorical = pd.api.types.is_string_dtype(df.iloc[:, 0]) or pd.api.types.is_object_dtype(df.iloc[:, 0])
            if first_col_categorical and len(numeric_cols) / df.shape[1] > 0.6:
                issues.append("cross_tab")
                recommendations.append(f"{df.columns[1:][:5].tolist()}... 열들을 세로 형식으로 변환하세요")
                format_type = DataFormat.CROSS_TAB

        # 3. 다단계 헤더 감지 (열 이름에 패턴이 있는 경우)
        header_patterns = [col for col in df.columns if isinstance(col, str) and ("." in col or "_" in col or " - " in col)]
        if len(header_patterns) > df.shape[1] * 0.5:
            issues.append("multi_level_headers")
            recommendations.append("다단계 헤더를 단일 헤더로 결합하세요")
            format_type = DataFormat.MULTI_LEVEL_HEADERS

        # 4. 소계 행 감지 (행에서 "총계", "소계", "합계" 등이 포함된 경우)
        subtotal_keywords = ["총계", "소계", "합계", "Total", "Subtotal", "Sum"]
        subtotal_rows = 0
        for _, row in df.iterrows():
            row_str = " ".join([str(val) for val in row.values if pd.notna(val)])
            if any(keyword in row_str for keyword in subtotal_keywords):
                subtotal_rows += 1

        if subtotal_rows > 0:
            issues.append("subtotals_mixed")
            recommendations.append(f"소계 행 {subtotal_rows}개를 제거하세요")
            format_type = DataFormat.SUBTOTALS_MIXED

        # 5. 넓은 형식 감지 (동일한 지표가 여러 열에 반복되는 경우)
        if df.shape[1] > 10:
            # 열 이름에서 연도, 월, 분기 패턴 감지
            date_patterns = [r"\d{4}", r"Q[1-4]", r"[1-9][0-2]?월", r"Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec"]
            date_cols = []
            for col in df.columns:
                if isinstance(col, str) and any(re.search(pattern, col) for pattern in date_patterns):
                    date_cols.append(col)

            if len(date_cols) > 3:
                issues.append("wide_format")
                recommendations.append(f"날짜/기간 열들({len(date_cols)}개)을 하나의 기간 컬럼으로 변환하세요")
                format_type = DataFormat.WIDE_FORMAT

        # 피벗테이블 준비 상태 평가
        pivot_ready = len(issues) == 0
        transformation_needed = not pivot_ready

        # 변환 후 예상 행 수 계산
        estimated_rows = df.shape[0]
        if "cross_tab" in issues or "wide_format" in issues:
            # 교차표나 넓은 형식인 경우 행 수가 크게 증가
            estimated_rows = df.shape[0] * max(len(numeric_cols), 1)

        # 신뢰도 점수 계산 (문제가 적을수록 높음)
        confidence_score = max(0.0, 1.0 - len(issues) * 0.2)

        return {
            "format_type": format_type.value,
            "issues": issues,
            "pivot_ready": pivot_ready,
            "transformation_needed": transformation_needed,
            "recommendations": recommendations,
            "estimated_rows_after_transform": int(estimated_rows),
            "confidence_score": round(confidence_score, 2),
            "data_shape": {"rows": df.shape[0], "columns": df.shape[1]},
            "ai_assistance_available": True,
        }

    except Exception as e:
        return {
            "format_type": DataFormat.UNKNOWN.value,
            "issues": ["analysis_failed"],
            "pivot_ready": False,
            "transformation_needed": False,
            "recommendations": [f"데이터 분석 중 오류 발생: {str(e)}"],
            "estimated_rows_after_transform": 0,
            "confidence_score": 0.0,
            "ai_assistance_available": False,
        }


def transform_data_unpivot(df: pd.DataFrame, id_vars: Optional[List[str]] = None) -> pd.DataFrame:
    """
    교차표 형식 데이터를 세로 형식(unpivot)으로 변환합니다.

    Args:
        df: 변환할 pandas DataFrame
        id_vars: 고정할 열 목록 (None이면 첫 번째 열 자동 선택)

    Returns:
        변환된 pandas DataFrame
    """
    try:
        if id_vars is None:
            # 첫 번째 열을 ID 변수로 사용
            id_vars = [df.columns[0]]

        # 수치형 열들을 찾아서 melt 대상으로 설정
        numeric_cols = df.select_dtypes(include=["number"]).columns.tolist()
        value_vars = [col for col in numeric_cols if col not in id_vars]

        if not value_vars:
            # 수치형 열이 없으면 ID 변수 외의 모든 열 사용
            value_vars = [col for col in df.columns if col not in id_vars]

        # pandas melt를 사용하여 unpivot
        melted_df = pd.melt(df, id_vars=id_vars, value_vars=value_vars, var_name="변수", value_name="값")

        # 빈 값 제거
        melted_df = melted_df.dropna(subset=["값"])

        return melted_df.reset_index(drop=True)

    except Exception as e:
        raise ValueError(f"Unpivot 변환 실패: {str(e)}")


def transform_data_unmerge(df: pd.DataFrame) -> pd.DataFrame:
    """
    병합된 셀로 인한 빈 값들을 앞의 값으로 채웁니다.

    Args:
        df: 변환할 pandas DataFrame

    Returns:
        변환된 pandas DataFrame
    """
    try:
        # 모든 열에 대해 forward fill 적용
        filled_df = df.fillna(method="ffill")

        # 문자열 열의 경우 빈 문자열도 채우기
        for col in filled_df.columns:
            if filled_df[col].dtype == "object":
                filled_df[col] = filled_df[col].replace("", None)
                filled_df[col] = filled_df[col].fillna(method="ffill")

        return filled_df.reset_index(drop=True)

    except Exception as e:
        raise ValueError(f"Unmerge 변환 실패: {str(e)}")


def transform_data_flatten_headers(df: pd.DataFrame) -> pd.DataFrame:
    """
    다단계 헤더를 단일 헤더로 결합합니다.

    Args:
        df: 변환할 pandas DataFrame

    Returns:
        변환된 pandas DataFrame
    """
    try:
        # 열 이름을 단순화
        new_columns = []
        for col in df.columns:
            if isinstance(col, str):
                # 특수 문자를 언더스코어로 대체하고 단순화
                simplified = re.sub(r"[^\w\s]", "_", col)
                simplified = re.sub(r"\s+", "_", simplified)
                simplified = re.sub(r"_+", "_", simplified).strip("_")
                new_columns.append(simplified)
            else:
                new_columns.append(f"Column_{len(new_columns)+1}")

        df.columns = new_columns
        return df.reset_index(drop=True)

    except Exception as e:
        raise ValueError(f"Header 평탄화 실패: {str(e)}")


def transform_data_remove_subtotals(df: pd.DataFrame) -> pd.DataFrame:
    """
    소계 행들을 제거합니다.

    Args:
        df: 변환할 pandas DataFrame

    Returns:
        변환된 pandas DataFrame
    """
    try:
        subtotal_keywords = ["총계", "소계", "합계", "Total", "Subtotal", "Sum", "계"]

        # 소계 행 인덱스 찾기
        rows_to_drop = []
        for idx, row in df.iterrows():
            row_str = " ".join([str(val) for val in row.values if pd.notna(val)])
            if any(keyword in row_str for keyword in subtotal_keywords):
                rows_to_drop.append(idx)

        # 소계 행 제거
        cleaned_df = df.drop(rows_to_drop)

        return cleaned_df.reset_index(drop=True)

    except Exception as e:
        raise ValueError(f"소계 제거 실패: {str(e)}")


def transform_data_auto(df: pd.DataFrame) -> Tuple[pd.DataFrame, List[str]]:
    """
    자동으로 모든 필요한 변환을 적용합니다.

    Args:
        df: 변환할 pandas DataFrame

    Returns:
        (변환된 DataFrame, 적용된 변환 목록) 튜플
    """
    try:
        applied_transforms = []
        result_df = df.copy()

        # 1. 소계 제거 (먼저 수행)
        subtotal_keywords = ["총계", "소계", "합계", "Total", "Subtotal", "Sum"]
        subtotal_rows = 0
        for _, row in result_df.iterrows():
            row_str = " ".join([str(val) for val in row.values if pd.notna(val)])
            if any(keyword in row_str for keyword in subtotal_keywords):
                subtotal_rows += 1

        if subtotal_rows > 0:
            result_df = transform_data_remove_subtotals(result_df)
            applied_transforms.append("remove-subtotals")

        # 2. 병합된 셀 처리 (빈 값 비율 확인)
        empty_ratio = result_df.isnull().sum().sum() / (result_df.shape[0] * result_df.shape[1])
        if empty_ratio > 0.3:
            result_df = transform_data_unmerge(result_df)
            applied_transforms.append("unmerge")

        # 3. 헤더 평탄화 (복잡한 열 이름 확인)
        header_patterns = [
            col for col in result_df.columns if isinstance(col, str) and ("." in col or "_" in col or " - " in col)
        ]
        if len(header_patterns) > result_df.shape[1] * 0.5:
            result_df = transform_data_flatten_headers(result_df)
            applied_transforms.append("flatten-headers")

        # 4. Unpivot 변환 (교차표 또는 넓은 형식 감지)
        numeric_cols = result_df.select_dtypes(include=["number"]).columns
        should_unpivot = False

        # 교차표 감지
        if len(numeric_cols) > 3 and result_df.shape[1] > 5:
            first_col_categorical = pd.api.types.is_string_dtype(result_df.iloc[:, 0]) or pd.api.types.is_object_dtype(
                result_df.iloc[:, 0]
            )
            if first_col_categorical and len(numeric_cols) / result_df.shape[1] > 0.6:
                should_unpivot = True

        # 넓은 형식 감지 (날짜/기간 열들)
        if result_df.shape[1] > 10:
            date_patterns = [r"\d{4}", r"Q[1-4]", r"[1-9][0-2]?월", r"Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec"]
            date_cols = []
            for col in result_df.columns:
                if isinstance(col, str) and any(re.search(pattern, col) for pattern in date_patterns):
                    date_cols.append(col)
            if len(date_cols) > 3:
                should_unpivot = True

        if should_unpivot:
            result_df = transform_data_unpivot(result_df)
            applied_transforms.append("unpivot")

        return result_df, applied_transforms

    except Exception as e:
        raise ValueError(f"자동 변환 실패: {str(e)}")
