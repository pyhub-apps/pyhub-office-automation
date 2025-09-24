#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HeadVer 버전 태그 생성 스크립트

HeadVer 형식: v{major}.{yearweek}.{build}
- major: 메이저 버전 (.headver 파일에서 읽음)
- yearweek: 년도 뒤 2자리 + ISO 주차 2자리 (예: 2539 = 2025년 39주차)
- build: 빌드 번호 (수동 입력 또는 자동 증가)

사용법:
    python scripts/create_version_tag.py [build_number] [--message "Release message"]

예시:
    python scripts/create_version_tag.py 18
    python scripts/create_version_tag.py 19 --message "Fix critical bug"
    python scripts/create_version_tag.py --auto-increment
"""

import argparse
import datetime
import subprocess
import sys
from pathlib import Path

# Windows 콘솔 UTF-8 출력 설정
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass  # 설정 실패해도 계속 진행


def get_major_version():
    """읽어온다 .headver 파일에서 메이저 버전을"""
    headver_file = Path(".headver")
    if not headver_file.exists():
        raise FileNotFoundError(".headver 파일이 존재하지 않습니다")

    with open(headver_file, "r") as f:
        major = f.read().strip()

    try:
        return int(major)
    except ValueError:
        raise ValueError(f".headver 파일의 내용이 숫자가 아닙니다: {major}")


def get_current_yearweek():
    """현재 년도 뒤 2자리 + ISO 주차를 YYWW 형식으로 반환"""
    now = datetime.datetime.now()
    year = now.year % 100  # 년도 뒤 2자리
    week = now.isocalendar()[1]  # ISO week number
    return f"{year}{week:02d}"


def get_latest_build_number(major, yearweek):
    """현재 year-week에 대한 최신 빌드 번호를 Git 태그에서 찾음"""
    try:
        # git tag -l "v{major}.{yearweek}.*" 형식으로 검색
        result = subprocess.run(["git", "tag", "-l", f"v{major}.{yearweek}.*"], capture_output=True, text=True, check=True)

        tags = result.stdout.strip().split("\n") if result.stdout.strip() else []

        # 빌드 번호 추출 및 최대값 찾기
        build_numbers = []
        for tag in tags:
            if tag:  # 빈 문자열 제외
                try:
                    # v10.2539.18 -> 18 추출
                    build_part = tag.split(".")[-1]
                    build_numbers.append(int(build_part))
                except (IndexError, ValueError):
                    continue

        return max(build_numbers) if build_numbers else 0

    except subprocess.CalledProcessError:
        return 0


def create_version_tag(major, yearweek, build, message=None):
    """버전 태그 생성"""
    tag_name = f"v{major}.{yearweek}.{build}"

    # 기본 메시지 생성
    if not message:
        message = f"Release {tag_name}"

    try:
        # 태그 생성
        subprocess.run(["git", "tag", "-a", tag_name, "-m", message], check=True)

        print(f"[SUCCESS] 태그 생성 완료: {tag_name}")
        print(f"[MESSAGE] {message}")

        # 원격으로 푸시 여부 확인
        response = input(f"\n원격 저장소에 태그를 푸시하시겠습니까? (y/N): ").strip().lower()
        if response in ["y", "yes"]:
            subprocess.run(["git", "push", "origin", tag_name], check=True)
            print(f"[SUCCESS] 원격 푸시 완료: {tag_name}")
            print(f"[INFO] GitHub Actions: https://github.com/pyhub-apps/pyhub-office-automation/actions")
        else:
            print(f"[WARNING] 로컬에만 태그가 생성되었습니다. 수동으로 푸시하려면:")
            print(f"   git push origin {tag_name}")

        return tag_name

    except subprocess.CalledProcessError as e:
        print(f"[ERROR] 태그 생성 실패: {e}")
        return None


def main():
    parser = argparse.ArgumentParser(
        description="HeadVer 형식의 Git 태그 생성",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
예시:
  python scripts/create_version_tag.py 18
  python scripts/create_version_tag.py 19 --message "Fix critical security bug"
  python scripts/create_version_tag.py --auto-increment --message "New feature release"
        """,
    )

    parser.add_argument("build", nargs="?", type=int, help="빌드 번호 (생략시 --auto-increment 필요)")

    parser.add_argument("--auto-increment", action="store_true", help="최신 빌드 번호에서 자동으로 1 증가")

    parser.add_argument("--message", "-m", type=str, help='태그 메시지 (기본값: "Release v{version}")')

    parser.add_argument("--dry-run", action="store_true", help="실제 태그를 만들지 않고 미리보기만")

    args = parser.parse_args()

    try:
        # 메이저 버전 읽기
        major = get_major_version()
        yearweek = get_current_yearweek()

        print(f"[INFO] 현재 버전 정보:")
        print(f"   메이저 버전: {major}")
        print(f"   년-주차: {yearweek} ({datetime.datetime.now().year}년 {datetime.datetime.now().isocalendar()[1]}주차)")

        # 빌드 번호 결정
        if args.auto_increment:
            latest_build = get_latest_build_number(major, yearweek)
            build = latest_build + 1
            print(f"   최신 빌드: {latest_build} -> 새 빌드: {build}")
        elif args.build is not None:
            build = args.build
            print(f"   지정된 빌드: {build}")
        else:
            parser.error("빌드 번호를 지정하거나 --auto-increment 옵션을 사용하세요")

        # 버전 태그 미리보기
        tag_name = f"v{major}.{yearweek}.{build}"
        print(f"\n[PREVIEW] 생성될 태그: {tag_name}")

        if args.dry_run:
            print("[DRY-RUN] 실제 태그를 생성하지 않습니다")
            return

        # 태그 생성
        result = create_version_tag(major, yearweek, build, args.message)
        if result:
            print(f"\n[COMPLETE] 성공적으로 완료되었습니다!")

    except Exception as e:
        print(f"[ERROR] 오류 발생: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
