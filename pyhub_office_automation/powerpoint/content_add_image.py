"""
PowerPoint 이미지 추가 명령어
슬라이드에 이미지를 추가합니다.
"""

import json
from pathlib import Path
from typing import Optional

import typer
from PIL import Image
from pptx import Presentation
from pptx.util import Inches

from pyhub_office_automation.version import get_version

from .utils import (
    calculate_aspect_ratio_size,
    create_error_response,
    create_success_response,
    normalize_path,
    validate_slide_number,
)


def content_add_image(
    file_path: str = typer.Option(..., "--file-path", help="PowerPoint 파일 경로"),
    slide_number: int = typer.Option(..., "--slide-number", help="이미지를 추가할 슬라이드 번호 (1부터 시작)"),
    image_path: str = typer.Option(..., "--image-path", help="추가할 이미지 파일 경로"),
    left: Optional[float] = typer.Option(None, "--left", help="이미지 왼쪽 위치 (인치)"),
    top: Optional[float] = typer.Option(None, "--top", help="이미지 상단 위치 (인치)"),
    width: Optional[float] = typer.Option(None, "--width", help="이미지 너비 (인치) - 미지정시 원본 비율 유지"),
    height: Optional[float] = typer.Option(None, "--height", help="이미지 높이 (인치) - 미지정시 원본 비율 유지"),
    center: bool = typer.Option(False, "--center", help="슬라이드 중앙에 배치 (--left, --top 무시)"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 (json/text)"),
):
    """
    PowerPoint 슬라이드에 이미지를 추가합니다.

    위치 지정 방법:
      --center: 슬라이드 중앙에 배치
      --left, --top: 특정 위치에 배치

    크기 지정:
      --width, --height: 둘 다 지정하면 지정된 크기로
      --width만 지정: 너비 기준으로 비율 유지하여 높이 자동 계산
      --height만 지정: 높이 기준으로 비율 유지하여 너비 자동 계산
      미지정: 원본 크기 (DPI 기준 인치 변환)

    예제:
        oa ppt content-add-image --file-path "presentation.pptx" --slide-number 1 --image-path "logo.png" --center --width 2
        oa ppt content-add-image --file-path "presentation.pptx" --slide-number 2 --image-path "chart.jpg" --left 1 --top 2 --height 3
        oa ppt content-add-image --file-path "presentation.pptx" --slide-number 3 --image-path "photo.png" --center
    """
    try:
        # 입력 검증
        if not center and (left is None or top is None):
            raise ValueError("--center를 사용하지 않는 경우 --left와 --top을 모두 지정해야 합니다")

        # 파일 경로 정규화 및 존재 확인
        normalized_pptx_path = normalize_path(file_path)
        pptx_path = Path(normalized_pptx_path).resolve()

        if not pptx_path.exists():
            raise FileNotFoundError(f"PowerPoint 파일을 찾을 수 없습니다: {pptx_path}")

        normalized_image_path = normalize_path(image_path)
        img_path = Path(normalized_image_path).resolve()

        if not img_path.exists():
            raise FileNotFoundError(f"이미지 파일을 찾을 수 없습니다: {img_path}")

        # 이미지 정보 읽기 (PIL 사용)
        with Image.open(str(img_path)) as img:
            original_width_px, original_height_px = img.size
            dpi = img.info.get("dpi", (96, 96))
            if isinstance(dpi, tuple):
                dpi_x, dpi_y = dpi
            else:
                dpi_x = dpi_y = dpi

            # 픽셀을 인치로 변환
            original_width_in = original_width_px / dpi_x
            original_height_in = original_height_px / dpi_y

        # 크기 계산 (aspect ratio 유지)
        if width is None and height is None:
            # 둘 다 미지정: 원본 크기 사용
            final_width = original_width_in
            final_height = original_height_in
        elif width is not None and height is not None:
            # 둘 다 지정: 지정된 크기 사용 (비율 무시)
            final_width = width
            final_height = height
        elif width is not None:
            # 너비만 지정: 높이를 비율에 맞춰 계산
            final_width, final_height = calculate_aspect_ratio_size(original_width_in, original_height_in, target_width=width)
        else:
            # 높이만 지정: 너비를 비율에 맞춰 계산
            final_width, final_height = calculate_aspect_ratio_size(
                original_width_in, original_height_in, target_height=height
            )

        # 프레젠테이션 열기
        prs = Presentation(str(pptx_path))

        # 슬라이드 번호 검증
        slide_idx = validate_slide_number(slide_number, len(prs.slides))
        slide = prs.slides[slide_idx]

        # 위치 계산
        if center:
            # 슬라이드 크기 가져오기 (EMU 단위)
            slide_width = prs.slide_width
            slide_height = prs.slide_height

            # 인치 단위로 변환
            slide_width_in = slide_width / 914400  # 1 inch = 914400 EMU
            slide_height_in = slide_height / 914400

            # 중앙 배치 위치 계산
            final_left = (slide_width_in - final_width) / 2
            final_top = (slide_height_in - final_height) / 2
        else:
            final_left = left
            final_top = top

        # 이미지 추가
        picture = slide.shapes.add_picture(
            str(img_path), Inches(final_left), Inches(final_top), width=Inches(final_width), height=Inches(final_height)
        )

        # 저장
        prs.save(str(pptx_path))

        # 결과 데이터 구성
        result_data = {
            "file": str(pptx_path),
            "slide_number": slide_number,
            "image_file": str(img_path),
            "position": {
                "left": round(final_left, 2),
                "top": round(final_top, 2),
                "width": round(final_width, 2),
                "height": round(final_height, 2),
            },
            "original_size": {
                "width_px": original_width_px,
                "height_px": original_height_px,
                "width_in": round(original_width_in, 2),
                "height_in": round(original_height_in, 2),
            },
            "centered": center,
        }

        # 성공 응답
        message = f"슬라이드 {slide_number}에 이미지를 추가했습니다"
        if center:
            message += " (중앙 배치)"

        response = create_success_response(
            data=result_data,
            command="content-add-image",
            message=message,
        )

        # 출력
        if output_format == "json":
            typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
        else:
            typer.echo(f"✅ {message}")
            typer.echo(f"📄 파일: {pptx_path.name}")
            typer.echo(f"📍 슬라이드: {slide_number}")
            typer.echo(f"🖼️ 이미지: {img_path.name}")
            typer.echo(f"📐 위치: {result_data['position']['left']}in × {result_data['position']['top']}in")
            typer.echo(f"📏 크기: {result_data['position']['width']}in × {result_data['position']['height']}in")
            typer.echo(
                f"🎨 원본: {result_data['original_size']['width_px']}px × {result_data['original_size']['height_px']}px"
            )

    except FileNotFoundError as e:
        error_response = create_error_response(e, "content-add-image")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)

    except ValueError as e:
        error_response = create_error_response(e, "content-add-image")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)

    except Exception as e:
        error_response = create_error_response(e, "content-add-image")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ 예기치 않은 오류: {str(e)}", err=True)
        raise typer.Exit(1)


if __name__ == "__main__":
    typer.run(content_add_image)
