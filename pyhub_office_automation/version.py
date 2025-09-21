"""
HeadVer 버전 관리 시스템
Format: {head}.{yearweek}.{build}
Reference: https://github.com/line/headver
"""

import os
import subprocess
from datetime import datetime
from pathlib import Path


def get_head_version():
    """Head 버전을 .headver 파일에서 읽어옴"""
    headver_file = Path(__file__).parent.parent / ".headver"
    if headver_file.exists():
        return headver_file.read_text().strip()
    return "1"


def get_yearweek():
    """현재 날짜의 yearweek 계산 (YYWW 형식)"""
    now = datetime.now()
    year = now.strftime("%y")
    week = now.strftime("%V")
    return f"{year}{week:0>2}"


def get_build_number():
    """빌드 번호 계산"""
    # GitHub Actions 환경에서는 GITHUB_RUN_NUMBER 사용
    if "GITHUB_RUN_NUMBER" in os.environ:
        return os.environ["GITHUB_RUN_NUMBER"]

    # CI/CD BUILD_NUMBER 환경변수 사용
    if "BUILD_NUMBER" in os.environ:
        return os.environ["BUILD_NUMBER"]

    # 로컬 개발 환경에서는 git commit count 사용
    try:
        # 이번 주 첫날(월요일)부터의 커밋 수 계산
        result = subprocess.run(
            ["git", "rev-list", "--count", "--since=last monday", "HEAD"], capture_output=True, text=True, check=True
        )
        return result.stdout.strip() or "0"
    except (subprocess.CalledProcessError, FileNotFoundError):
        return "0"


def get_version():
    """HeadVer 형식의 전체 버전 문자열 생성"""
    head = get_head_version()
    yearweek = get_yearweek()
    build = get_build_number()
    return f"{head}.{yearweek}.{build}"


def get_version_info():
    """버전 정보를 딕셔너리로 반환"""
    head = get_head_version()
    yearweek = get_yearweek()
    build = get_build_number()
    version = f"{head}.{yearweek}.{build}"

    return {"head": head, "yearweek": yearweek, "build": build, "version": version}


if __name__ == "__main__":
    print(get_version())
