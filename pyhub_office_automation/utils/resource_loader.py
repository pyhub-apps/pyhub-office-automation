"""
PyInstaller 호환 리소스 파일 로더
- 개발 환경과 패키징된 환경 모두 지원
- 안전한 파일 경로 처리
"""

import os
import sys
from pathlib import Path
from typing import Optional, Union


def get_resource_path(resource_name: str) -> Optional[Path]:
    """
    리소스 파일의 경로를 반환합니다.
    PyInstaller 환경과 개발 환경 모두 지원합니다.

    Args:
        resource_name (str): 리소스 파일명 (예: "welcome.md")

    Returns:
        Optional[Path]: 리소스 파일 경로, 파일이 없으면 None
    """
    # PyInstaller 환경 확인
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        # PyInstaller로 패키징된 환경
        # sys._MEIPASS는 임시 디렉토리를 가리킴
        base_path = Path(sys._MEIPASS)
        resource_path = base_path / "pyhub_office_automation" / "resources" / resource_name
    else:
        # 개발 환경
        # 현재 파일의 위치를 기준으로 상대 경로 계산
        current_file = Path(__file__)
        # utils -> pyhub_office_automation -> resources
        base_path = current_file.parent.parent
        resource_path = base_path / "resources" / resource_name

    # 파일 존재 여부 확인
    if resource_path.exists() and resource_path.is_file():
        return resource_path
    else:
        return None


def load_resource_text(resource_name: str, encoding: str = "utf-8") -> Optional[str]:
    """
    리소스 파일의 텍스트 내용을 로드합니다.

    Args:
        resource_name (str): 리소스 파일명 (예: "welcome.md")
        encoding (str): 파일 인코딩 (기본값: utf-8)

    Returns:
        Optional[str]: 파일 내용, 파일을 읽을 수 없으면 None
    """
    resource_path = get_resource_path(resource_name)

    if resource_path is None:
        return None

    try:
        with open(resource_path, "r", encoding=encoding) as f:
            return f.read()
    except (OSError, IOError, UnicodeDecodeError) as e:
        # 파일 읽기 실패 시 None 반환
        print(f"Warning: Failed to load resource '{resource_name}': {e}")
        return None


def get_fallback_message(message_type: str) -> str:
    """
    리소스 파일을 로드할 수 없을 때 사용할 폴백 메시지를 반환합니다.

    Args:
        message_type (str): 메시지 타입 ("welcome" 또는 "llm-guide")

    Returns:
        str: 폴백 메시지
    """
    if message_type == "welcome":
        return """
🎉 pyhub-office-automation에 오신 것을 환영합니다!

이 도구는 AI 에이전트를 위한 Office 자동화 도구입니다.

빠른 시작:
• oa info                    - 설치 상태 확인
• oa excel list              - Excel 명령어 목록
• oa excel workbook-list     - 열린 워크북 확인
• oa install-guide           - 설치 가이드

도움말:
• oa --help                  - 전체 명령어 도움말
• oa llm-guide               - AI 에이전트 사용 지침
"""
    elif message_type == "llm-guide":
        return """
# LLM/AI 에이전트를 위한 사용 가이드

## 핵심 명령어
- oa info: 패키지 정보 확인
- oa excel list --format json: Excel 명령어 목록
- oa excel workbook-list: 현재 열린 워크북 확인

## 연결 방법
1. --file-path: 파일 경로로 연결
2. --use-active: 활성 워크북 사용
3. --workbook-name: 워크북 이름으로 연결

## 에러 방지
작업 전 항상 workbook-list로 상황 파악하세요.

자세한 내용은 oa --help를 참조하세요.
"""
    else:
        return f"Unknown message type: {message_type}"


def load_welcome_message() -> str:
    """Welcome 메시지를 로드합니다."""
    content = load_resource_text("welcome.md")
    return content if content is not None else get_fallback_message("welcome")


def load_llm_guide() -> str:
    """LLM 가이드를 로드합니다."""
    content = load_resource_text("llm-guide.md")
    return content if content is not None else get_fallback_message("llm-guide")
