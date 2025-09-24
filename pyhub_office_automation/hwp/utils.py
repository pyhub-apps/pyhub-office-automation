"""
HWP 자동화를 위한 공통 유틸리티 함수들
pyhwpx 라이브러리를 활용한 HWP 문서 처리 도구
"""

import json
import os
import platform
import tempfile
import time
from pathlib import Path
from typing import Any, Dict, Optional, Union

import typer

from pyhub_office_automation.version import get_version


def normalize_path(path: str) -> str:
    """
    파일 경로 정규화 (macOS 한글 처리 포함)

    Args:
        path: 정규화할 파일 경로

    Returns:
        정규화된 절대 경로
    """
    if not path:
        return path

    # macOS에서 한글 파일명 정규화 (NFC 형태로 변환)
    if platform.system() == "Darwin":
        import unicodedata
        path = unicodedata.normalize('NFC', path)

    # 절대 경로로 변환
    return str(Path(path).resolve())


def check_hwp_installed() -> bool:
    """
    HWP(한글) 프로그램 설치 여부 확인

    Returns:
        HWP 설치 여부
    """
    if platform.system() != "Windows":
        return False

    try:
        # pyhwpx import with warning suppression (COM 캐시 재구축 경고 방지)
        import warnings
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            import pyhwpx
        # 실제로 HWP COM 객체 생성 테스트
        hwp = pyhwpx.Hwp()
        hwp.quit()
        return True
    except Exception:
        return False


def create_success_response(
    data: Any = None,
    processing_stats: Optional[Dict] = None,
    metadata: Optional[Dict] = None,
    command: str = "export"
) -> Dict[str, Any]:
    """
    성공 응답 생성

    Args:
        data: 반환할 데이터
        processing_stats: 처리 통계 정보
        metadata: 메타데이터
        command: 실행된 명령어명

    Returns:
        구조화된 성공 응답
    """
    response = {
        "command": command,
        "version": get_version(),
        "status": "success"
    }

    if data is not None:
        response["data"] = data

    if processing_stats:
        response["processing_stats"] = processing_stats

    if metadata:
        response["metadata"] = metadata

    return response


def create_error_response(
    error_message: str,
    error_type: str = "ClickException",
    command: str = "export"
) -> Dict[str, Any]:
    """
    에러 응답 생성

    Args:
        error_message: 에러 메시지
        error_type: 에러 타입
        command: 실행된 명령어명

    Returns:
        구조화된 에러 응답
    """
    return {
        "command": command,
        "version": get_version(),
        "status": "error",
        "error_message": error_message,
        "error_type": error_type
    }


def format_output(data: Dict[str, Any], output_format: str = "json") -> str:
    """
    출력 형식에 맞게 데이터 포맷팅

    Args:
        data: 출력할 데이터
        output_format: 출력 형식 (json, yaml 등)

    Returns:
        포맷팅된 문자열
    """
    if output_format.lower() == "json":
        try:
            return json.dumps(data, ensure_ascii=False, indent=2)
        except UnicodeEncodeError:
            return json.dumps(data, ensure_ascii=True, indent=2)
    else:
        # 기본적으로 JSON 반환
        return json.dumps(data, ensure_ascii=False, indent=2)


def cleanup_temp_file(file_path: str) -> bool:
    """
    임시 파일 정리

    Args:
        file_path: 정리할 파일 경로

    Returns:
        정리 성공 여부
    """
    try:
        if os.path.exists(file_path):
            os.remove(file_path)
            return True
        return True  # 파일이 없으면 성공으로 간주
    except Exception:
        return False


def clean_html_content(html_content: str) -> str:
    """
    HWP에서 생성된 HTML에서 불필요한 CSS/메타데이터 제거

    Args:
        html_content: 원본 HTML 내용

    Returns:
        정리된 HTML 내용
    """
    import re

    # HWP 특유의 불필요한 CSS 클래스 제거
    patterns_to_remove = [
        r'<meta[^>]*generator[^>]*>',  # Generator 메타 태그
        r'class="[^"]*HwpObj[^"]*"',   # HWP 객체 클래스
        r'style="[^"]*position:\s*absolute[^"]*"',  # 절대 위치 스타일
        r'<!\-\-[^>]*\-\->',  # HTML 주석
    ]

    cleaned_content = html_content
    for pattern in patterns_to_remove:
        cleaned_content = re.sub(pattern, '', cleaned_content, flags=re.IGNORECASE)

    # 불필요한 공백 정리
    cleaned_content = re.sub(r'\n\s*\n', '\n', cleaned_content)
    cleaned_content = re.sub(r'>\s+<', '><', cleaned_content)

    return cleaned_content.strip()


class ExecutionTimer:
    """실행 시간 측정을 위한 컨텍스트 매니저"""

    def __init__(self):
        self.start_time = None
        self.end_time = None

    def __enter__(self):
        self.start_time = time.time()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.end_time = time.time()

    @property
    def duration_ms(self) -> int:
        """실행 시간을 밀리초 단위로 반환"""
        if self.start_time and self.end_time:
            return int((self.end_time - self.start_time) * 1000)
        return 0


def validate_file_path(file_path: str) -> str:
    """
    파일 경로 유효성 검증 및 정규화

    Args:
        file_path: 검증할 파일 경로

    Returns:
        정규화된 파일 경로

    Raises:
        typer.BadParameter: 파일이 존재하지 않거나 HWP 파일이 아닌 경우
    """
    if not file_path:
        raise typer.BadParameter("파일 경로를 지정해야 합니다")

    normalized_path = normalize_path(file_path)

    if not os.path.exists(normalized_path):
        raise typer.BadParameter(f"파일이 존재하지 않습니다: {normalized_path}")

    if not normalized_path.lower().endswith('.hwp'):
        raise typer.BadParameter(f"HWP 파일만 지원됩니다: {normalized_path}")

    return normalized_path


def get_file_size(file_path: str) -> int:
    """
    파일 크기 반환 (바이트 단위)

    Args:
        file_path: 파일 경로

    Returns:
        파일 크기 (바이트)
    """
    try:
        return os.path.getsize(file_path)
    except Exception:
        return 0


def create_temp_html_file() -> str:
    """
    임시 HTML 파일 경로 생성

    Returns:
        임시 HTML 파일 경로
    """
    return tempfile.mktemp(suffix='.html', prefix='hwp_export_')