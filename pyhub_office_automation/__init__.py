"""
pyhub-office-automation: Python 기반 Excel 및 HWP 자동화 패키지

AI 에이전트(주로 Gemini CLI)를 위한 대화형 office 문서 자동화 도구
"""

from .version import get_version

# 동적 버전 가져오기
_version = get_version()
__version__ = _version
__author__ = "pyhub-apps"
__email__ = "admin@pyhub.kr"
__description__ = "Python-based Excel and HWP automation package for AI agents"
