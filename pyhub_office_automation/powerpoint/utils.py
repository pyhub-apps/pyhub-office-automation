"""
PowerPoint 자동화를 위한 공통 유틸리티 함수들
python-pptx 기반 PowerPoint 조작 및 데이터 처리 지원
"""

import json
import platform
import sys
import unicodedata
from enum import Enum
from pathlib import Path
from typing import Any, Dict, Optional, Union

from pyhub_office_automation.version import get_version


# CLI 명령어 인자를 위한 Enum 클래스들
class OutputFormat(str, Enum):
    """출력 형식 선택지"""

    JSON = "json"
    TEXT = "text"
    MARKDOWN = "markdown"


class LayoutType(str, Enum):
    """PowerPoint 레이아웃 타입"""

    TITLE = "title"
    TITLE_AND_CONTENT = "title_and_content"
    SECTION_HEADER = "section_header"
    TWO_CONTENT = "two_content"
    COMPARISON = "comparison"
    TITLE_ONLY = "title_only"
    BLANK = "blank"
    CONTENT_WITH_CAPTION = "content_with_caption"
    PICTURE_WITH_CAPTION = "picture_with_caption"


class ExportFormat(str, Enum):
    """프레젠테이션 내보내기 형식"""

    PDF = "pdf"
    PNG = "png"
    JPG = "jpg"
    JPEG = "jpeg"
    GIF = "gif"


class PowerPointBackend(str, Enum):
    """PowerPoint 백엔드 선택지"""

    PYTHON_PPTX = "python-pptx"
    COM = "com"
    AUTO = "auto"


def get_powerpoint_backend() -> str:
    """
    현재 플랫폼에서 사용 가능한 PowerPoint 백엔드를 반환합니다.

    Returns:
        str: 'python-pptx' (크로스플랫폼) 또는 'com' (Windows 전용)
    """
    if platform.system() == "Windows":
        # Windows에서는 COM 사용 가능
        try:
            import win32com.client

            return PowerPointBackend.COM.value
        except ImportError:
            pass

    # 모든 플랫폼에서 python-pptx 사용 가능
    try:
        import pptx

        return PowerPointBackend.PYTHON_PPTX.value
    except ImportError:
        raise ImportError("python-pptx 패키지가 설치되지 않았습니다. 'pip install python-pptx'로 설치하세요.")


def check_feature_availability(feature: str) -> Dict[str, Any]:
    """
    특정 기능의 플랫폼별 가용성을 체크합니다.

    Args:
        feature: 체크할 기능 이름

    Returns:
        Dict: 기능 가용성 정보
            - available: 현재 플랫폼에서 사용 가능 여부
            - backend: 사용할 백엔드 ('python-pptx' or 'com')
            - platform: 현재 플랫폼
            - limitations: 제한사항 목록
    """
    current_platform = platform.system()
    backend = get_powerpoint_backend()

    # 기본 기능 (모든 플랫폼 지원)
    basic_features = [
        "presentation-create",
        "presentation-open",
        "presentation-save",
        "slide-add",
        "slide-delete",
        "content-add-text",
        "content-add-image",
    ]

    # COM 전용 기능 (Windows만)
    com_only_features = ["macro-run", "animation-add", "transition-set"]

    limitations = []
    available = True

    if feature in com_only_features:
        if current_platform != "Windows":
            available = False
            limitations.append(f"{feature}는 Windows에서만 사용 가능합니다")
        elif backend != PowerPointBackend.COM.value:
            available = False
            limitations.append(f"{feature}는 Windows COM (pywin32)이 필요합니다")

    return {
        "available": available,
        "backend": backend,
        "platform": current_platform,
        "feature": feature,
        "limitations": limitations,
    }


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


def create_success_response(
    command: str,
    data: Any,
    message: Optional[str] = None,
    version: Optional[str] = None,
) -> Dict[str, Any]:
    """
    성공 응답을 표준 JSON 형식으로 생성합니다.

    Args:
        command: 실행한 명령어 이름
        data: 반환할 데이터
        message: 추가 메시지 (선택사항)
        version: 패키지 버전 (기본값: 현재 버전)

    Returns:
        Dict: 표준화된 성공 응답
    """
    response = {
        "success": True,
        "command": command,
        "data": data,
        "version": version or get_version(),
    }

    if message:
        response["message"] = message

    return response


def create_error_response(
    command: str,
    error: Union[str, Exception],
    error_type: Optional[str] = None,
    version: Optional[str] = None,
) -> Dict[str, Any]:
    """
    에러 응답을 표준 JSON 형식으로 생성합니다.

    Args:
        command: 실행한 명령어 이름
        error: 에러 메시지 또는 Exception 객체
        error_type: 에러 타입 (선택사항)
        version: 패키지 버전 (기본값: 현재 버전)

    Returns:
        Dict: 표준화된 에러 응답
    """
    error_message = str(error)
    if isinstance(error, Exception):
        error_type = error_type or type(error).__name__

    response = {
        "success": False,
        "command": command,
        "error": error_message,
        "error_type": error_type or "UnknownError",
        "version": version or get_version(),
    }

    return response


def get_or_open_presentation(
    file_path: Optional[str] = None,
    presentation_name: Optional[str] = None,
    backend: str = "auto",
) -> Any:
    """
    파일 경로 또는 프레젠테이션 이름으로 프레젠테이션을 열거나 가져옵니다.

    Args:
        file_path: 프레젠테이션 파일 경로
        presentation_name: 열려있는 프레젠테이션 이름
        backend: 사용할 백엔드 ('auto', 'python-pptx', 'com')

    Returns:
        Presentation 객체 (python-pptx) 또는 COM 객체

    Raises:
        ValueError: 파일 경로와 프레젠테이션 이름이 모두 없거나 둘 다 제공된 경우
        FileNotFoundError: 파일이 존재하지 않는 경우
        ImportError: 필요한 라이브러리가 설치되지 않은 경우
    """
    # 입력 검증
    if file_path and presentation_name:
        raise ValueError("file_path와 presentation_name 중 하나만 지정해야 합니다")

    if not file_path and not presentation_name:
        raise ValueError("file_path 또는 presentation_name 중 하나는 필수입니다")

    # 백엔드 선택
    if backend == "auto":
        backend = get_powerpoint_backend()

    # 경로 정규화
    if file_path:
        file_path = normalize_path(file_path)
        if not Path(file_path).exists():
            raise FileNotFoundError(f"프레젠테이션 파일을 찾을 수 없습니다: {file_path}")

    # python-pptx 백엔드
    if backend == PowerPointBackend.PYTHON_PPTX.value:
        try:
            from pptx import Presentation

            if file_path:
                return Presentation(file_path)
            else:
                raise NotImplementedError(
                    "python-pptx는 프레젠테이션 이름으로 열기를 지원하지 않습니다. file_path를 사용하세요."
                )

        except ImportError:
            raise ImportError("python-pptx 패키지가 설치되지 않았습니다. 'pip install python-pptx'로 설치하세요.")

    # COM 백엔드 (Windows 전용)
    elif backend == PowerPointBackend.COM.value:
        if platform.system() != "Windows":
            raise NotImplementedError("COM 백엔드는 Windows에서만 사용 가능합니다")

        try:
            import win32com.client

            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            powerpoint.Visible = True

            if file_path:
                return powerpoint.Presentations.Open(str(Path(file_path).resolve()))
            else:
                # 이름으로 찾기
                for pres in powerpoint.Presentations:
                    if pres.Name == presentation_name:
                        return pres
                raise FileNotFoundError(f"열려있는 프레젠테이션을 찾을 수 없습니다: {presentation_name}")

        except ImportError:
            raise ImportError("pywin32 패키지가 설치되지 않았습니다. 'pip install pywin32'로 설치하세요.")

    else:
        raise ValueError(f"지원하지 않는 백엔드입니다: {backend}")


def output_result(result: Dict[str, Any], output_format: str = "json") -> None:
    """
    결과를 지정된 형식으로 출력합니다.

    Args:
        result: 출력할 결과 딕셔너리
        output_format: 출력 형식 ('json', 'text', 'markdown')
    """
    if output_format == OutputFormat.JSON.value:
        try:
            # UTF-8 인코딩으로 JSON 출력
            json_output = json.dumps(result, ensure_ascii=False, indent=2)
            print(json_output)
        except UnicodeEncodeError:
            # 인코딩 실패 시 ASCII로 폴백
            json_output = json.dumps(result, ensure_ascii=True, indent=2)
            print(json_output)

    elif output_format == OutputFormat.TEXT.value:
        # 텍스트 형식 출력
        if result.get("success"):
            print(f"✓ {result.get('command')} - Success")
            if result.get("message"):
                print(f"  {result['message']}")
        else:
            print(f"✗ {result.get('command')} - Error")
            print(f"  {result.get('error')}")

    elif output_format == OutputFormat.MARKDOWN.value:
        # 마크다운 형식 출력
        print(f"## {result.get('command')}")
        print()
        if result.get("success"):
            print("**Status**: ✓ Success")
            if result.get("message"):
                print(f"**Message**: {result['message']}")
        else:
            print("**Status**: ✗ Error")
            print(f"**Error**: {result.get('error')}")
        print()


def validate_output_format(format_str: str) -> str:
    """
    출력 형식 문자열을 검증합니다.

    Args:
        format_str: 검증할 형식 문자열

    Returns:
        검증된 형식 문자열

    Raises:
        ValueError: 지원하지 않는 형식인 경우
    """
    format_lower = format_str.lower()
    valid_formats = [f.value for f in OutputFormat]

    if format_lower not in valid_formats:
        raise ValueError(f"지원하지 않는 출력 형식입니다: {format_str}. 사용 가능: {', '.join(valid_formats)}")

    return format_lower


# Slide 관리 유틸리티 함수들 (Issue #76)


def get_layout_by_name_or_index(prs, identifier: Union[str, int]):
    """
    레이아웃을 이름(str) 또는 인덱스(int)로 찾습니다.

    Args:
        prs: Presentation 객체
        identifier: 레이아웃 이름(문자열) 또는 인덱스(정수)

    Returns:
        SlideLayout 객체

    Raises:
        ValueError: 레이아웃을 찾을 수 없는 경우
        IndexError: 인덱스가 범위를 벗어난 경우
    """
    # 인덱스로 찾기
    if isinstance(identifier, int):
        try:
            return prs.slide_layouts[identifier]
        except IndexError:
            raise IndexError(f"레이아웃 인덱스 범위 초과: {identifier} (사용 가능: 0-{len(prs.slide_layouts)-1})")

    # 이름으로 찾기
    if isinstance(identifier, str):
        for layout in prs.slide_layouts:
            if layout.name == identifier:
                return layout

        # 사용 가능한 레이아웃 목록 제공
        available_layouts = [f"{i}: {layout.name}" for i, layout in enumerate(prs.slide_layouts)]
        raise ValueError(
            f"레이아웃을 찾을 수 없습니다: '{identifier}'\n" f"사용 가능한 레이아웃:\n" + "\n".join(available_layouts)
        )

    raise TypeError(f"레이아웃 식별자는 문자열 또는 정수여야 합니다: {type(identifier)}")


def validate_slide_number(slide_num: int, total_slides: int, allow_append: bool = False) -> int:
    """
    슬라이드 번호를 검증하고 0-based 인덱스로 변환합니다.

    Args:
        slide_num: 검증할 슬라이드 번호 (1-based)
        total_slides: 총 슬라이드 수
        allow_append: True면 total_slides+1도 허용 (끝에 추가 시)

    Returns:
        int: 0-based 인덱스

    Raises:
        TypeError: 슬라이드 번호가 정수가 아닌 경우
        ValueError: 슬라이드 번호가 범위를 벗어난 경우
    """
    if not isinstance(slide_num, int):
        raise TypeError(f"슬라이드 번호는 정수여야 합니다: {slide_num} ({type(slide_num).__name__})")

    max_value = total_slides + 1 if allow_append else total_slides

    if total_slides == 0:
        # 빈 프레젠테이션인 경우
        if allow_append and slide_num == 1:
            return 0
        raise ValueError("프레젠테이션에 슬라이드가 없습니다")

    if not (1 <= slide_num <= max_value):
        raise ValueError(f"슬라이드 번호 범위: 1-{max_value}, 입력값: {slide_num}")

    return slide_num - 1  # 0-based 인덱스 반환


def get_slide_content_summary(slide) -> Dict[str, int]:
    """
    슬라이드의 콘텐츠 요약을 생성합니다 (도형 타입별 개수).

    Args:
        slide: Slide 객체

    Returns:
        Dict[str, int]: 도형 타입별 개수
            - textbox: 텍스트 박스 수
            - picture: 이미지 수
            - chart: 차트 수
            - table: 표 수
            - other: 기타 도형 수

    Example:
        >>> summary = get_slide_content_summary(slide)
        >>> print(summary)
        {"textbox": 3, "picture": 1, "chart": 0, "table": 1, "other": 0}
    """
    try:
        from pptx.enum.shapes import MSO_SHAPE_TYPE
    except ImportError:
        # python-pptx 없으면 빈 요약 반환
        return {"textbox": 0, "picture": 0, "chart": 0, "table": 0, "other": 0}

    summary = {"textbox": 0, "picture": 0, "chart": 0, "table": 0, "other": 0}

    for shape in slide.shapes:
        try:
            if shape.shape_type == MSO_SHAPE_TYPE.TEXT_BOX:
                summary["textbox"] += 1
            elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                summary["picture"] += 1
            elif shape.shape_type == MSO_SHAPE_TYPE.CHART:
                summary["chart"] += 1
            elif shape.shape_type == MSO_SHAPE_TYPE.TABLE:
                summary["table"] += 1
            else:
                summary["other"] += 1
        except Exception:
            # 도형 타입 확인 실패 시 other로 카운트
            summary["other"] += 1

    return summary


def get_slide_title(slide) -> str:
    """
    슬라이드의 제목을 추출합니다.

    Args:
        slide: Slide 객체

    Returns:
        str: 슬라이드 제목 (없으면 빈 문자열)

    Example:
        >>> title = get_slide_title(slide)
        >>> print(title)
        "Q4 Sales Report"
    """
    try:
        if hasattr(slide, "shapes") and hasattr(slide.shapes, "title"):
            title_shape = slide.shapes.title
            if title_shape and hasattr(title_shape, "text"):
                return title_shape.text.strip()
    except Exception:
        # 제목 추출 실패 시 빈 문자열 반환
        pass

    return ""
