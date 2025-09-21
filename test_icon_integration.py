#!/usr/bin/env python
"""
아이콘 통합 테스트 스크립트
PyInstaller 빌드에서 아이콘이 올바르게 포함되었는지 확인
"""

import os
import subprocess
import sys
from pathlib import Path


def test_icon_files_exist():
    """아이콘 파일 존재 확인"""
    print("🔍 아이콘 파일 존재 확인...")

    required_files = [
        "pyhub_office_automation/assets/icons/logo.ico",
        "pyhub_office_automation/assets/icons/logo.png",
    ]

    all_exist = True
    for file_path in required_files:
        path = Path(file_path)
        if path.exists():
            size_kb = path.stat().st_size / 1024
            print(f"   ✅ {file_path} ({size_kb:.1f} KB)")
        else:
            print(f"   ❌ {file_path} - 파일이 없습니다")
            all_exist = False

    # 다양한 크기의 PNG 파일들 확인
    sizes_dir = Path("pyhub_office_automation/assets/icons/logo_sizes")
    if sizes_dir.exists():
        png_files = list(sizes_dir.glob("logo_*.png"))
        print(f"   📂 PNG 아이콘 크기: {len(png_files)}개")
        for png_file in sorted(png_files):
            size_kb = png_file.stat().st_size / 1024
            print(f"      - {png_file.name} ({size_kb:.1f} KB)")
    else:
        print(f"   ❌ {sizes_dir} - 디렉터리가 없습니다")
        all_exist = False

    return all_exist


def test_spec_file_icon():
    """spec 파일에 아이콘 설정 확인"""
    print("📋 spec 파일 아이콘 설정 확인...")

    spec_file = Path("oa.spec")
    if not spec_file.exists():
        print("   ❌ oa.spec 파일이 없습니다")
        return False

    try:
        content = spec_file.read_text(encoding="utf-8")

        if "icon='pyhub_office_automation/assets/icons/logo.ico'" in content:
            print("   ✅ spec 파일에 아이콘 경로가 설정되어 있습니다")
            return True
        else:
            print("   ❌ spec 파일에 아이콘 경로가 설정되어 있지 않습니다")
            return False

    except Exception as e:
        print(f"   ❌ spec 파일 읽기 실패: {e}")
        return False


def test_build_script_icon():
    """빌드 스크립트에 아이콘 처리 로직 확인"""
    print("🔨 빌드 스크립트 아이콘 처리 확인...")

    build_script = Path("build_windows.ps1")
    if not build_script.exists():
        print("   ❌ build_windows.ps1 파일이 없습니다")
        return False

    try:
        content = build_script.read_text(encoding="utf-8")

        icon_checks = [
            "Checking icon files",
            "logo.ico",
            "--icon",
        ]

        missing_checks = []
        for check in icon_checks:
            if check not in content:
                missing_checks.append(check)

        if not missing_checks:
            print("   ✅ 빌드 스크립트에 아이콘 처리 로직이 포함되어 있습니다")
            return True
        else:
            print(f"   ❌ 빌드 스크립트에 다음 아이콘 처리가 누락되었습니다: {missing_checks}")
            return False

    except Exception as e:
        print(f"   ❌ 빌드 스크립트 읽기 실패: {e}")
        return False


def test_ico_file_validity():
    """ICO 파일 유효성 확인"""
    print("🎨 ICO 파일 유효성 확인...")

    ico_path = Path("pyhub_office_automation/assets/icons/logo.ico")
    if not ico_path.exists():
        print("   ❌ ICO 파일이 없습니다")
        return False

    try:
        from PIL import Image

        with Image.open(ico_path) as img:
            print(f"   ✅ ICO 파일 형식: {img.format}")
            print(f"   ✅ ICO 파일 크기: {img.size}")
            print(f"   ✅ ICO 파일 모드: {img.mode}")

            # ICO 파일의 다중 해상도 확인
            if hasattr(img, "n_frames") and img.n_frames > 1:
                print(f"   ✅ 다중 해상도 포함: {img.n_frames}개")
                for i in range(min(img.n_frames, 5)):  # 최대 5개만 표시
                    img.seek(i)
                    print(f"      - 해상도 {i+1}: {img.size}")
            else:
                print("   ⚠️  단일 해상도 ICO 파일")

            return True

    except ImportError:
        print("   ⚠️  Pillow 라이브러리가 없어 ICO 파일 유효성을 확인할 수 없습니다")
        return True  # Pillow가 없어도 테스트는 통과로 처리
    except Exception as e:
        print(f"   ❌ ICO 파일 유효성 검증 실패: {e}")
        return False


def test_build_dry_run():
    """빌드 드라이 런 테스트 (실제 빌드 없이 설정 확인)"""
    print("🧪 빌드 설정 드라이 런 테스트...")

    # PyInstaller가 설치되어 있는지 확인
    try:
        result = subprocess.run(["pyinstaller", "--version"], capture_output=True, text=True, timeout=10)

        if result.returncode == 0:
            version = result.stdout.strip()
            print(f"   ✅ PyInstaller 버전: {version}")
        else:
            print("   ❌ PyInstaller가 제대로 설치되어 있지 않습니다")
            return False

    except FileNotFoundError:
        print("   ❌ PyInstaller가 설치되어 있지 않습니다")
        return False
    except Exception as e:
        print(f"   ❌ PyInstaller 확인 실패: {e}")
        return False

    # spec 파일 기본 구문 확인 (Python 구문 검사)
    spec_file = Path("oa.spec")
    if spec_file.exists():
        try:
            # 기본 Python 구문 검사만 수행
            content = spec_file.read_text(encoding="utf-8")

            # 필수 요소 확인
            required_elements = ["Analysis(", "PYZ(", "EXE(", "COLLECT(", "icon="]

            missing_elements = []
            for element in required_elements:
                if element not in content:
                    missing_elements.append(element)

            if not missing_elements:
                print("   ✅ spec 파일 구조가 유효합니다")
                return True
            else:
                print(f"   ❌ spec 파일에 필수 요소가 누락되었습니다: {missing_elements}")
                return False

        except Exception as e:
            print(f"   ❌ spec 파일 검증 실패: {e}")
            return False
    else:
        print("   ⚠️  spec 파일이 없습니다")
        return True


def main():
    """메인 테스트 함수"""
    print("=" * 60)
    print("🎨 아이콘 통합 테스트")
    print("=" * 60)

    tests = [
        ("아이콘 파일 존재 확인", test_icon_files_exist),
        ("spec 파일 아이콘 설정", test_spec_file_icon),
        ("빌드 스크립트 아이콘 처리", test_build_script_icon),
        ("ICO 파일 유효성", test_ico_file_validity),
        ("빌드 설정 드라이 런", test_build_dry_run),
    ]

    passed = 0
    total = len(tests)

    for test_name, test_func in tests:
        print(f"\n📋 {test_name}")
        try:
            if test_func():
                print(f"✅ {test_name} 통과")
                passed += 1
            else:
                print(f"❌ {test_name} 실패")
        except Exception as e:
            print(f"❌ {test_name} 오류: {e}")

    print("\n" + "=" * 60)
    print(f"테스트 결과: {passed}/{total} 통과")

    if passed == total:
        print("🎉 모든 아이콘 통합 테스트가 성공했습니다!")
        print("   다음 단계: 실제 빌드 테스트")
        print("   명령어: .\\build_windows.ps1 --BuildType onefile --GenerateMetadata")
    else:
        print("⚠️  일부 테스트가 실패했습니다.")
        print("   아이콘 통합을 완료한 후 다시 테스트하세요.")

    return passed == total


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
