#!/usr/bin/env python
"""
COM 캐시 재구축 경고 수정 테스트 스크립트
PyInstaller 빌드 후 첫 실행 시 COM 경고 메시지가 나타나는지 확인
"""

import os
import subprocess
import sys
import tempfile
from pathlib import Path


def test_excel_command():
    """Excel 명령어 테스트 - COM 초기화 확인"""
    print("🧪 Testing Excel command for COM warnings...")

    # dist 디렉터리에서 실행 파일 찾기
    possible_paths = ["dist/oa/oa.exe", "dist/oa.exe"]

    exe_path = None
    for path in possible_paths:
        if os.path.exists(path):
            exe_path = path
            break

    if not exe_path:
        print("❌ oa.exe 실행 파일을 찾을 수 없습니다.")
        print("   빌드가 완료되었는지 확인하세요.")
        return False

    print(f"   Found executable: {exe_path}")

    # 임시 파일 생성
    with tempfile.NamedTemporaryFile(mode="w", suffix=".txt", delete=False) as f:
        f.write("test data")
        temp_file = f.name

    try:
        # Excel 명령어 실행 (도움말)
        print("   Running: oa excel --help")
        result = subprocess.run([exe_path, "excel", "--help"], capture_output=True, text=True, timeout=30)

        print(f"   Exit code: {result.returncode}")

        # 출력에서 COM 관련 경고 메시지 확인
        stderr_output = result.stderr.lower()
        com_warnings = [
            "rebuilding cache of generated files for com support",
            "could not add module",
            "circular import",
            "_get_good_object_",
        ]

        warnings_found = []
        for warning in com_warnings:
            if warning in stderr_output:
                warnings_found.append(warning)

        if warnings_found:
            print("❌ COM 경고 메시지가 발견되었습니다:")
            for warning in warnings_found:
                print(f"   - {warning}")
            return False
        else:
            print("✅ COM 경고 메시지가 발견되지 않았습니다.")
            return True

    except subprocess.TimeoutExpired:
        print("❌ 명령어 실행 시간 초과")
        return False
    except Exception as e:
        print(f"❌ 테스트 실행 중 오류: {e}")
        return False
    finally:
        # 임시 파일 정리
        try:
            os.unlink(temp_file)
        except:
            pass


def test_version_command():
    """버전 명령어 테스트 - 기본 기능 확인"""
    print("🧪 Testing version command...")

    # dist 디렉터리에서 실행 파일 찾기
    possible_paths = ["dist/oa/oa.exe", "dist/oa.exe"]

    exe_path = None
    for path in possible_paths:
        if os.path.exists(path):
            exe_path = path
            break

    if not exe_path:
        return False

    try:
        result = subprocess.run([exe_path, "--version"], capture_output=True, text=True, timeout=10)

        print(f"   Exit code: {result.returncode}")
        print(f"   Output: {result.stdout.strip()}")

        return result.returncode == 0

    except Exception as e:
        print(f"❌ 버전 테스트 실행 중 오류: {e}")
        return False


def main():
    """메인 테스트 함수"""
    print("=" * 50)
    print("COM 캐시 재구축 경고 수정 테스트")
    print("=" * 50)

    # 빌드 파일 존재 확인
    if not any(os.path.exists(p) for p in ["dist/oa/oa.exe", "dist/oa.exe"]):
        print("❌ 빌드된 실행 파일을 찾을 수 없습니다.")
        print("   먼저 빌드를 실행하세요: .\\build_windows.ps1")
        return False

    # 테스트 실행
    tests = [
        ("버전 명령어", test_version_command),
        ("Excel 명령어 (COM 경고 확인)", test_excel_command),
    ]

    passed = 0
    total = len(tests)

    for test_name, test_func in tests:
        print(f"\n📋 {test_name}")
        if test_func():
            print(f"✅ {test_name} 통과")
            passed += 1
        else:
            print(f"❌ {test_name} 실패")

    print("\n" + "=" * 50)
    print(f"테스트 결과: {passed}/{total} 통과")

    if passed == total:
        print("🎉 모든 테스트가 성공했습니다!")
        print("   COM 캐시 재구축 경고가 해결되었습니다.")
    else:
        print("⚠️  일부 테스트가 실패했습니다.")
        print("   추가 조치가 필요할 수 있습니다.")

    return passed == total


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
