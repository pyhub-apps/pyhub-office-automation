"""
PowerPoint 백엔드 자동 선택 로직
플랫폼과 환경에 따라 최적의 백엔드를 선택
"""

import platform
from typing import Optional, Tuple

from .utils import PowerPointBackend


def detect_backend(force_backend: Optional[str] = None) -> str:
    """
    사용 가능한 PowerPoint 백엔드를 감지하고 선택합니다.

    우선순위:
    1. force_backend가 지정된 경우: 강제로 해당 백엔드 사용
    2. Windows + pywin32 설치: COM 백엔드 (완전한 기능)
    3. 기타 환경: python-pptx 백엔드 (제한적 기능)

    Args:
        force_backend: 강제로 사용할 백엔드 ("com", "python-pptx", None)

    Returns:
        선택된 백엔드 문자열 ("com" 또는 "python-pptx")

    Raises:
        ValueError: 지원하지 않는 백엔드가 지정된 경우
        RuntimeError: 강제 지정된 백엔드를 사용할 수 없는 경우

    Example:
        >>> backend = detect_backend()
        >>> print(f"Selected backend: {backend}")
        Selected backend: com  # Windows + pywin32
    """
    # 강제 지정된 백엔드가 있는 경우
    if force_backend:
        backend_lower = force_backend.lower()

        # 유효성 검증
        valid_backends = [PowerPointBackend.COM.value, PowerPointBackend.PYTHON_PPTX.value]
        if backend_lower not in valid_backends:
            raise ValueError(
                f"지원하지 않는 백엔드입니다: {force_backend}\n" f"사용 가능한 백엔드: {', '.join(valid_backends)}"
            )

        # COM 강제 사용 시 환경 체크
        if backend_lower == PowerPointBackend.COM.value:
            if platform.system() != "Windows":
                raise RuntimeError("COM 백엔드는 Windows에서만 사용 가능합니다\n" f"현재 플랫폼: {platform.system()}")

            try:
                import win32com.client
            except ImportError:
                raise RuntimeError("COM 백엔드를 사용하려면 pywin32 패키지가 필요합니다\n" "설치: pip install pywin32")

        # python-pptx 강제 사용 시 환경 체크
        elif backend_lower == PowerPointBackend.PYTHON_PPTX.value:
            try:
                import pptx
            except ImportError:
                raise RuntimeError(
                    "python-pptx 백엔드를 사용하려면 python-pptx 패키지가 필요합니다\n" "설치: pip install python-pptx"
                )

        return backend_lower

    # 자동 선택: Windows + pywin32가 있으면 COM 우선
    if platform.system() == "Windows":
        try:
            import win32com.client

            return PowerPointBackend.COM.value
        except ImportError:
            pass

    # 기본: python-pptx
    try:
        import pptx

        return PowerPointBackend.PYTHON_PPTX.value
    except ImportError:
        raise RuntimeError(
            "사용 가능한 PowerPoint 백엔드가 없습니다\n"
            "다음 중 하나를 설치하세요:\n"
            "- Windows COM: pip install pywin32\n"
            "- 크로스플랫폼: pip install python-pptx"
        )


def check_backend_availability(backend: str) -> Tuple[bool, Optional[str]]:
    """
    지정된 백엔드가 사용 가능한지 체크합니다.

    Args:
        backend: 체크할 백엔드 ("com" 또는 "python-pptx")

    Returns:
        (available: bool, error_message: Optional[str]) 튜플
        - available: 사용 가능 여부
        - error_message: 사용 불가능한 경우 에러 메시지 (사용 가능하면 None)

    Example:
        >>> available, error = check_backend_availability("com")
        >>> if not available:
        ...     print(f"COM 백엔드 사용 불가: {error}")
    """
    backend_lower = backend.lower()

    # COM 백엔드 체크
    if backend_lower == PowerPointBackend.COM.value:
        if platform.system() != "Windows":
            return False, f"COM 백엔드는 Windows 전용입니다 (현재: {platform.system()})"

        try:
            import win32com.client

            return True, None
        except ImportError:
            return False, "pywin32 패키지가 설치되지 않았습니다 (pip install pywin32)"

    # python-pptx 백엔드 체크
    elif backend_lower == PowerPointBackend.PYTHON_PPTX.value:
        try:
            import pptx

            return True, None
        except ImportError:
            return False, "python-pptx 패키지가 설치되지 않았습니다 (pip install python-pptx)"

    else:
        return False, f"알 수 없는 백엔드입니다: {backend}"


def get_backend_info(backend: str) -> dict:
    """
    백엔드 정보를 반환합니다.

    Args:
        backend: 백엔드 이름 ("com" 또는 "python-pptx")

    Returns:
        백엔드 정보 딕셔너리
        {
            "backend": "com",
            "name": "Windows COM",
            "platform": "Windows",
            "features": ["full"],
            "limitations": []
        }

    Example:
        >>> info = get_backend_info("com")
        >>> print(f"Backend: {info['name']}")
        Backend: Windows COM
    """
    backend_lower = backend.lower()

    if backend_lower == PowerPointBackend.COM.value:
        return {
            "backend": PowerPointBackend.COM.value,
            "name": "Windows COM (pywin32)",
            "platform": "Windows",
            "features": [
                "완전한 PowerPoint 제어",
                "레이아웃 변경",
                "테마 적용",
                "애니메이션/전환효과",
                "매크로 실행",
                "슬라이드쇼 제어",
                "VBA 수준의 자동화",
            ],
            "limitations": ["Windows 전용"],
            "recommended": True,
        }

    elif backend_lower == PowerPointBackend.PYTHON_PPTX.value:
        return {
            "backend": PowerPointBackend.PYTHON_PPTX.value,
            "name": "python-pptx (크로스플랫폼)",
            "platform": "Windows, macOS, Linux",
            "features": ["기본 프레젠테이션 생성/수정", "슬라이드 추가/삭제", "텍스트/이미지/표 삽입", "기본 차트 생성"],
            "limitations": [
                "⚠️ 레이아웃 변경 불가 (조회만)",
                "⚠️ 테마 적용 제한",
                "❌ 애니메이션/전환효과 미지원",
                "❌ 매크로 실행 불가",
                "❌ 슬라이드쇼 제어 불가",
            ],
            "recommended": False,
        }

    else:
        return {
            "backend": backend,
            "name": "Unknown",
            "platform": "Unknown",
            "features": [],
            "limitations": ["알 수 없는 백엔드"],
            "recommended": False,
        }


def suggest_alternative_backend(current_backend: str, feature: str) -> Optional[str]:
    """
    현재 백엔드에서 지원하지 않는 기능에 대해 대체 백엔드를 제안합니다.

    Args:
        current_backend: 현재 백엔드 ("com" 또는 "python-pptx")
        feature: 필요한 기능 ("layout-change", "animation", "macro" 등)

    Returns:
        대체 백엔드 이름 또는 None (대체 불가능)

    Example:
        >>> alt = suggest_alternative_backend("python-pptx", "animation")
        >>> if alt:
        ...     print(f"애니메이션 기능을 사용하려면 {alt} 백엔드를 사용하세요")
    """
    current_lower = current_backend.lower()

    # python-pptx에서 COM 전용 기능 요청 시
    if current_lower == PowerPointBackend.PYTHON_PPTX.value:
        com_only_features = [
            "layout-change",
            "layout-apply",
            "theme-apply",
            "theme-change",
            "animation",
            "animation-add",
            "transition",
            "transition-set",
            "macro",
            "macro-run",
            "slideshow",
            "slideshow-start",
            "slideshow-control",
        ]

        if feature.lower() in com_only_features:
            # Windows 환경인지 체크
            if platform.system() == "Windows":
                available, _ = check_backend_availability(PowerPointBackend.COM.value)
                if available:
                    return PowerPointBackend.COM.value

    # COM에서는 모든 기능 지원
    return None


def get_backend_warning_message(backend: str, feature: Optional[str] = None) -> Optional[str]:
    """
    백엔드 제약사항에 대한 경고 메시지를 생성합니다.

    Args:
        backend: 백엔드 이름
        feature: 사용하려는 기능 (선택)

    Returns:
        경고 메시지 또는 None (경고 불필요)

    Example:
        >>> warning = get_backend_warning_message("python-pptx", "animation")
        >>> if warning:
        ...     print(f"⚠️ {warning}")
    """
    backend_lower = backend.lower()

    if backend_lower == PowerPointBackend.PYTHON_PPTX.value:
        # 기본 경고 메시지
        base_warning = (
            "⚠️ python-pptx 모드: 제한적 기능만 사용 가능합니다.\n"
            "완전한 기능을 사용하려면 Windows에서 COM 백엔드를 사용하세요."
        )

        # 특정 기능에 대한 추가 메시지
        if feature:
            alt_backend = suggest_alternative_backend(backend, feature)
            if alt_backend:
                return (
                    f"⚠️ '{feature}' 기능은 python-pptx에서 지원하지 않습니다.\n"
                    f"이 기능을 사용하려면 {alt_backend} 백엔드를 사용하세요.\n"
                    f"예: --backend {alt_backend}"
                )

        return base_warning

    # COM 백엔드는 경고 불필요
    return None
