"""
Excel 자동화를 위한 공통 유틸리티 함수들
xlwings 기반 Excel 조작 및 데이터 처리 지원
"""

import json
import csv
import io
import tempfile
import os
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union
import xlwings as xw
from ..version import get_version


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


def create_success_response(data: Any, command: str, message: str = None) -> Dict[str, Any]:
    """
    표준 성공 응답을 생성합니다.

    Args:
        data: 응답 데이터
        command: 명령어 이름
        message: 성공 메시지

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

    return response