"""
Excel 자동화를 위한 공통 유틸리티 함수들
xlwings 기반 Excel 조작 및 데이터 처리 지원
"""

import json
import csv
import io
import tempfile
import os
import unicodedata
import platform
import time
import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union
import xlwings as xw
from ..version import get_version


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
        return unicodedata.normalize('NFC', path)

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
    if '!' in range_str:
        sheet_name, range_part = range_str.split('!', 1)
        return sheet_name, range_part
    else:
        return None, range_str


def get_range(sheet: xw.Sheet, range_str: str, expand_mode: Optional[str] = None) -> xw.Range:
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
        if expand_mode == "table":
            range_obj = range_obj.expand()
        elif expand_mode == "down":
            range_obj = range_obj.expand('down')
        elif expand_mode == "right":
            range_obj = range_obj.expand('right')

    return range_obj


def handle_temp_file(data: Any, file_format: str = "json") -> str:
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

    temp_file = tempfile.NamedTemporaryFile(
        mode=mode,
        suffix=suffix,
        delete=False,
        encoding='utf-8'
    )

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


def format_output(data: Any, output_format: str = "json", include_version: bool = True) -> str:
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


def load_data_from_file(file_path: str) -> Any:
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
            with open(file_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        elif suffix == ".csv":
            with open(file_path, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                return [row for row in reader]
        else:
            with open(file_path, 'r', encoding='utf-8') as f:
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
    if '!' in range_str:
        _, range_part = range_str.split('!', 1)
    else:
        range_part = range_str

    # A1:B10 형태의 범위 패턴
    range_pattern = r'^[A-Z]+\d+(:[A-Z]+\d+)?$'
    return bool(re.match(range_pattern, range_part.upper()))


def create_error_response(error: Exception, command: str) -> Dict[str, Any]:
    """
    표준 에러 응답을 생성합니다.

    Args:
        error: 발생한 예외
        command: 명령어 이름

    Returns:
        에러 응답 딕셔너리
    """
    error_type = type(error).__name__

    response = {
        "success": False,
        "error_type": error_type,
        "error": str(error),
        "command": command,
        "version": get_version()
    }

    # 특정 에러에 대한 제안사항 추가
    if error_type == "FileNotFoundError":
        response["suggestion"] = "파일 경로를 확인하고 파일이 존재하는지 확인하세요."
    elif error_type == "RuntimeError" and "Excel" in str(error):
        response["suggestion"] = "Excel이 설치되어 있는지 확인하고, 다른 프로그램에서 파일을 사용 중이지 않은지 확인하세요."
    elif error_type == "ValueError" and "범위" in str(error):
        response["suggestion"] = "범위 형식이 올바른지 확인하세요. 예: 'A1:C10' 또는 'Sheet1!A1:C10'"

    return response


def create_success_response(
    data: Any,
    command: str,
    message: str = None,
    execution_time_ms: float = None,
    book: Optional[xw.Book] = None,
    **stats_kwargs
) -> Dict[str, Any]:
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
    response = {
        "success": True,
        "command": command,
        "version": get_version(),
        "data": data
    }

    if message:
        response["message"] = message

    # 메타데이터 추가
    metadata = {
        "command_category": COMMAND_CATEGORIES.get(command, "unknown"),
        "operation_type": OPERATION_TYPES.get(command, "unknown"),
        "timestamp": datetime.datetime.now().isoformat()
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
        response["current_context"] = {
            "total_open_workbooks": len(xw.books) if xw.books else 0,
            "collection_failed": True
        }

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
    file_path: Optional[str] = None,
    workbook_name: Optional[str] = None,
    use_active: bool = False,
    visible: bool = True
) -> xw.Book:
    """
    여러 방법으로 워크북을 가져오는 통합 함수입니다.

    Args:
        file_path: 파일 경로 (기존 방식)
        workbook_name: 열린 워크북 이름
        use_active: 활성 워크북 사용 여부
        visible: Excel 애플리케이션 표시 여부

    Returns:
        xlwings Book 객체

    Raises:
        ValueError: 옵션이 잘못 지정된 경우
        RuntimeError: 워크북을 찾거나 열 수 없는 경우
    """
    # 옵션 검증 - 정확히 하나만 지정되어야 함
    options_count = sum([
        bool(file_path),
        bool(workbook_name),
        use_active
    ])

    if options_count == 0:
        raise ValueError("file_path, workbook_name, use_active 중 하나는 반드시 지정해야 합니다")
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
        {"command": "range-read", "reason": "워크북의 데이터를 읽습니다"}
    ],
    "workbook-list": [
        {"command": "workbook-info", "reason": "특정 워크북의 상세 정보를 확인합니다"},
        {"command": "workbook-open", "reason": "새 워크북을 엽니다"}
    ],
    "workbook-info": [
        {"command": "range-read", "reason": "시트의 데이터를 확인합니다"},
        {"command": "sheet-activate", "reason": "특정 시트로 이동합니다"}
    ],
    "range-read": [
        {"command": "range-write", "reason": "데이터를 수정하거나 추가합니다"},
        {"command": "table-read", "reason": "테이블 형태로 더 많은 데이터를 읽습니다"}
    ],
    "range-write": [
        {"command": "range-read", "reason": "작성한 데이터를 확인합니다"},
        {"command": "workbook-save", "reason": "변경사항을 저장합니다"}
    ],
    "table-read": [
        {"command": "table-write", "reason": "테이블 데이터를 수정합니다"},
        {"command": "pivot-create", "reason": "피벗 테이블을 생성합니다"}
    ],
    "table-write": [
        {"command": "table-read", "reason": "작성한 테이블을 확인합니다"},
        {"command": "chart-add", "reason": "차트를 생성합니다"}
    ],
    "sheet-add": [
        {"command": "sheet-activate", "reason": "새 시트로 이동합니다"},
        {"command": "range-write", "reason": "새 시트에 데이터를 입력합니다"},
        {"command": "sheet-rename", "reason": "시트 이름을 변경합니다"}
    ],
    "sheet-activate": [
        {"command": "range-read", "reason": "활성 시트의 데이터를 확인합니다"},
        {"command": "sheet-add", "reason": "새 시트를 추가합니다"}
    ],
    "sheet-rename": [
        {"command": "sheet-activate", "reason": "이름이 변경된 시트로 이동합니다"},
        {"command": "workbook-info", "reason": "전체 시트 목록을 확인합니다"}
    ],
    "sheet-delete": [
        {"command": "workbook-info", "reason": "남은 시트들을 확인합니다"},
        {"command": "sheet-add", "reason": "새 시트를 추가합니다"}
    ],
    "pivot-create": [
        {"command": "pivot-configure", "reason": "피벗 테이블을 설정합니다"},
        {"command": "chart-pivot-create", "reason": "피벗 차트를 생성합니다"}
    ],
    "chart-add": [
        {"command": "chart-configure", "reason": "차트를 설정합니다"},
        {"command": "chart-position", "reason": "차트 위치를 조정합니다"}
    ]
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
    "chart-export": "chart"
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
    "chart-export": "read"
}


def get_execution_context(book: Optional[xw.Book] = None) -> Dict[str, Any]:
    """
    현재 Excel 실행 컨텍스트 정보를 수집합니다.

    Args:
        book: xlwings Book 객체 (None이면 활성 워크북 사용)

    Returns:
        컨텍스트 정보 딕셔너리
    """
    context = {
        "total_open_workbooks": len(xw.books) if xw.books else 0,
        "excel_app_visible": None,
        "current_workbook": None
    }

    try:
        if book is None and len(xw.books) > 0:
            book = xw.books.active

        if book:
            context.update({
                "current_workbook": {
                    "name": normalize_path(book.name),
                    "full_name": normalize_path(book.fullname),
                    "saved": getattr(book, 'saved', True),
                    "total_sheets": len(book.sheets),
                    "active_sheet": book.sheets.active.name if book.sheets else None,
                    "sheet_names": [sheet.name for sheet in book.sheets]
                },
                "excel_app_visible": book.app.visible if hasattr(book, 'app') else None
            })

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


def calculate_operation_stats(command: str, **kwargs) -> Dict[str, Any]:
    """
    작업별 통계 정보를 계산합니다.

    Args:
        command: 실행된 명령어
        **kwargs: 명령어별 통계 데이터

    Returns:
        작업 통계 딕셔너리
    """
    stats = {
        "command": command,
        "timestamp": datetime.datetime.now().isoformat()
    }

    # 명령어별 특화 통계
    if command in ["range-read", "range-write"]:
        if "range_obj" in kwargs:
            range_obj = kwargs["range_obj"]
            try:
                stats.update({
                    "cells_count": range_obj.count if hasattr(range_obj, 'count') else 1,
                    "rows_count": range_obj.rows.count if hasattr(range_obj, 'rows') else 1,
                    "columns_count": range_obj.columns.count if hasattr(range_obj, 'columns') else 1,
                    "range_address": range_obj.address if hasattr(range_obj, 'address') else None
                })
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


def suggest_next_commands(command: str, context: Dict[str, Any], **kwargs) -> List[Dict[str, str]]:
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
            suggestions.insert(0, {
                "command": "workbook-save",
                "reason": "워크북에 저장되지 않은 변경사항이 있습니다"
            })

        # 시트가 많은 경우 시트 관리 명령어 추천
        if workbook_info.get("total_sheets", 0) > 5:
            suggestions.append({
                "command": "workbook-info",
                "reason": "많은 시트가 있으므로 전체 구조를 확인하는 것이 좋습니다"
            })

    # 저장되지 않은 다른 워크북이 있는 경우
    if context.get("unsaved_workbooks"):
        suggestions.append({
            "command": "workbook-list",
            "reason": f"저장되지 않은 워크북이 {len(context['unsaved_workbooks'])}개 있습니다"
        })

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