"""
Excel 엔진 팩토리 및 공통 모듈

플랫폼에 맞는 Excel 엔진을 자동으로 선택하고 제공합니다.
"""

import platform
from typing import Optional

from .base import ChartInfo, ExcelEngineBase, RangeData, TableInfo, WorkbookInfo
from .exceptions import (
    AppleScriptError,
    ChartNotFoundError,
    COMError,
    DataValidationError,
    EngineInitializationError,
    ExcelEngineError,
    ExcelNotRunningError,
    PlatformNotSupportedError,
    RangeError,
    SheetNotFoundError,
    TableNotFoundError,
    WorkbookNotFoundError,
)

# 전역 엔진 인스턴스 (싱글톤)
_engine_instance: Optional[ExcelEngineBase] = None


def get_engine(force_platform: Optional[str] = None) -> ExcelEngineBase:
    """
    플랫폼에 맞는 Excel 엔진을 반환합니다 (싱글톤).

    Args:
        force_platform: 테스트용 플랫폼 강제 지정 ('windows', 'darwin')
                       None이면 자동 감지

    Returns:
        ExcelEngineBase: 플랫폼별 엔진 인스턴스

    Raises:
        PlatformNotSupportedError: 지원되지 않는 플랫폼인 경우
        EngineInitializationError: 엔진 초기화 실패

    Example:
        >>> engine = get_engine()
        >>> workbooks = engine.get_workbooks()
    """
    global _engine_instance

    if _engine_instance is not None:
        return _engine_instance

    # 플랫폼 감지
    current_platform = force_platform or platform.system()

    try:
        if current_platform == "Windows":
            # Windows: pywin32 COM 엔진
            from .windows import WindowsEngine

            _engine_instance = WindowsEngine()

        elif current_platform == "Darwin":  # macOS
            # macOS: AppleScript 엔진
            from .macos import MacOSEngine

            _engine_instance = MacOSEngine()

        else:
            raise PlatformNotSupportedError(current_platform)

    except ImportError as e:
        raise EngineInitializationError(current_platform, f"필요한 모듈을 가져올 수 없습니다: {str(e)}")

    except Exception as e:
        raise EngineInitializationError(current_platform, str(e))

    return _engine_instance


def reset_engine():
    """
    엔진 인스턴스를 초기화합니다 (테스트용).

    테스트 환경에서 플랫폼을 전환하거나 엔진을 재초기화할 때 사용합니다.

    Example:
        >>> reset_engine()
        >>> engine = get_engine(force_platform='Darwin')
    """
    global _engine_instance
    _engine_instance = None


def get_platform_name() -> str:
    """
    현재 플랫폼 이름을 반환합니다.

    Returns:
        str: 'Windows', 'macOS', 또는 'Unknown'
    """
    system = platform.system()
    if system == "Windows":
        return "Windows"
    elif system == "Darwin":
        return "macOS"
    else:
        return "Unknown"


# 공개 API
__all__ = [
    # Factory 함수
    "get_engine",
    "reset_engine",
    "get_platform_name",
    # 기본 클래스 및 데이터 클래스
    "ExcelEngineBase",
    "WorkbookInfo",
    "RangeData",
    "TableInfo",
    "ChartInfo",
    # 예외 클래스
    "ExcelEngineError",
    "WorkbookNotFoundError",
    "SheetNotFoundError",
    "RangeError",
    "TableNotFoundError",
    "ChartNotFoundError",
    "PlatformNotSupportedError",
    "ExcelNotRunningError",
    "COMError",
    "AppleScriptError",
    "DataValidationError",
    "EngineInitializationError",
]
