#!/usr/bin/env python
"""
아이콘 생성 스크립트
PNG 파일을 Windows ICO 형식으로 변환하고 다양한 크기의 PNG 아이콘 생성
"""

import os
import sys
from pathlib import Path
from typing import List, Tuple

try:
    from PIL import Image
except ImportError:
    print("❌ Pillow 라이브러리가 필요합니다.")
    print("   설치: pip install Pillow")
    sys.exit(1)


def create_resized_icons(source_path: Path, output_dir: Path, sizes: List[int]) -> List[Path]:
    """다양한 크기의 PNG 아이콘 생성"""
    print(f"📏 다양한 크기의 아이콘 생성: {sizes}")

    created_files = []

    try:
        with Image.open(source_path) as img:
            # 투명도 처리를 위해 RGBA로 변환
            if img.mode != "RGBA":
                img = img.convert("RGBA")

            for size in sizes:
                # 고품질 리샘플링으로 크기 조정
                resized = img.resize((size, size), Image.Resampling.LANCZOS)

                # 파일명 생성
                output_file = output_dir / f"logo_{size}.png"

                # PNG로 저장 (투명도 보존)
                resized.save(output_file, "PNG", optimize=True)
                created_files.append(output_file)

                print(f"   ✅ {size}x{size} → {output_file.name}")

    except Exception as e:
        print(f"❌ 아이콘 크기 조정 실패: {e}")
        return []

    return created_files


def create_ico_file(source_path: Path, output_path: Path, sizes: List[int]) -> bool:
    """다중 해상도 ICO 파일 생성"""
    print(f"🔄 ICO 파일 생성: {output_path.name}")

    try:
        with Image.open(source_path) as img:
            # 투명도 처리를 위해 RGBA로 변환
            if img.mode != "RGBA":
                img = img.convert("RGBA")

            # 각 크기별로 이미지 생성
            icon_images = []

            for size in sizes:
                # 고품질 리샘플링으로 크기 조정
                resized = img.resize((size, size), Image.Resampling.LANCZOS)
                icon_images.append(resized)
                print(f"   📐 {size}x{size} 해상도 추가")

            # ICO 파일로 저장 (다중 해상도)
            icon_images[0].save(
                output_path,
                format="ICO",
                sizes=[(size, size) for size in sizes],
                append_images=icon_images[1:] if len(icon_images) > 1 else None,
            )

            print(f"   ✅ ICO 파일 생성 완료: {output_path}")
            return True

    except Exception as e:
        print(f"❌ ICO 파일 생성 실패: {e}")
        return False


def validate_ico_file(ico_path: Path) -> bool:
    """ICO 파일 유효성 검증"""
    try:
        with Image.open(ico_path) as img:
            print(f"🔍 ICO 파일 검증: {ico_path.name}")
            print(f"   형식: {img.format}")
            print(f"   크기: {img.size}")
            print(f"   모드: {img.mode}")

            # ICO 파일의 모든 해상도 확인
            if hasattr(img, "n_frames"):
                print(f"   포함된 해상도 수: {img.n_frames}")

                for i in range(img.n_frames):
                    img.seek(i)
                    print(f"     - {i+1}: {img.size}")

            return True

    except Exception as e:
        print(f"❌ ICO 파일 검증 실패: {e}")
        return False


def main():
    """메인 실행 함수"""
    print("=" * 60)
    print("🎨 pyhub-office-automation 아이콘 생성기")
    print("=" * 60)

    # 경로 설정
    script_dir = Path(__file__).parent
    source_png = script_dir / "pyhub_office_automation" / "assets" / "icons" / "logo.png"
    icons_dir = source_png.parent
    sizes_dir = icons_dir / "logo_sizes"
    ico_file = icons_dir / "logo.ico"

    # 소스 파일 확인
    if not source_png.exists():
        print(f"❌ 소스 PNG 파일을 찾을 수 없습니다: {source_png}")
        return False

    print(f"📂 소스 파일: {source_png}")
    print(f"📁 출력 디렉터리: {icons_dir}")

    # 파일 정보 출력
    file_size = source_png.stat().st_size / (1024 * 1024)  # MB
    print(f"📊 파일 크기: {file_size:.2f} MB")

    try:
        with Image.open(source_png) as img:
            print(f"📐 원본 해상도: {img.size}")
            print(f"🎨 색상 모드: {img.mode}")
    except Exception as e:
        print(f"❌ 이미지 정보 읽기 실패: {e}")
        return False

    # Windows ICO 표준 크기
    ico_sizes = [256, 128, 48, 32, 16]
    png_sizes = [256, 128, 64, 48, 32, 16]

    print(f"\n🎯 생성할 ICO 크기: {ico_sizes}")
    print(f"🎯 생성할 PNG 크기: {png_sizes}")

    # 자동 진행 (CI/CD 환경 고려)
    print("\n✅ 자동으로 진행합니다.")

    print("\n" + "=" * 60)
    print("🚀 아이콘 생성 시작")
    print("=" * 60)

    # 1. 다양한 크기의 PNG 아이콘 생성
    png_files = create_resized_icons(source_png, sizes_dir, png_sizes)
    if not png_files:
        print("❌ PNG 아이콘 생성 실패")
        return False

    print(f"✅ {len(png_files)}개의 PNG 아이콘 생성 완료")

    # 2. ICO 파일 생성
    if not create_ico_file(source_png, ico_file, ico_sizes):
        print("❌ ICO 파일 생성 실패")
        return False

    print("✅ ICO 파일 생성 완료")

    # 3. ICO 파일 검증
    if not validate_ico_file(ico_file):
        print("❌ ICO 파일 검증 실패")
        return False

    print("✅ ICO 파일 검증 완료")

    # 결과 요약
    print("\n" + "=" * 60)
    print("🎉 아이콘 생성 완료!")
    print("=" * 60)
    print(f"📂 생성된 파일:")
    print(f"   ICO: {ico_file}")
    print(f"   PNG: {len(png_files)}개 ({sizes_dir})")

    # 파일 크기 정보
    ico_size = ico_file.stat().st_size / 1024  # KB
    print(f"📊 ICO 파일 크기: {ico_size:.1f} KB")

    print(f"\n📋 다음 단계:")
    print(f"   1. PyInstaller spec 파일에서 아이콘 경로 설정")
    print(f"   2. 빌드 스크립트에서 아이콘 검증 추가")
    print(f"   3. 테스트 빌드 실행")

    return True


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
