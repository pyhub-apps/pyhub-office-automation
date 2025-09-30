"""
PowerPoint 이미지 내보내기 명령어 (COM-First)
슬라이드를 이미지 파일로 변환하여 저장합니다.
"""

import json
from pathlib import Path
from typing import Optional

import typer

from pyhub_office_automation.version import get_version

from .utils import (
    PowerPointBackend,
    create_error_response,
    create_success_response,
    get_or_open_presentation,
    get_powerpoint_backend,
    normalize_path,
)


def export_images(
    output_dir: str = typer.Option(..., "--output-dir", help="이미지 저장 디렉토리"),
    image_format: str = typer.Option("PNG", "--format", help="이미지 형식 (PNG/JPG/GIF/BMP/TIFF)"),
    slides: Optional[str] = typer.Option(None, "--slides", help="내보낼 슬라이드 범위 (예: '1-5', '1,3,5', 'all')"),
    width: Optional[int] = typer.Option(None, "--width", help="이미지 너비 (픽셀)"),
    height: Optional[int] = typer.Option(None, "--height", help="이미지 높이 (픽셀)"),
    dpi: int = typer.Option(96, "--dpi", help="해상도 (DPI, 기본값: 96)"),
    filename_pattern: str = typer.Option("slide_{num:03d}", "--filename-pattern", help="파일명 패턴 (예: 'slide_{num:03d}')"),
    file_path: Optional[str] = typer.Option(None, "--file-path", help="PowerPoint 파일 경로"),
    presentation_name: Optional[str] = typer.Option(None, "--presentation-name", help="열려있는 프레젠테이션 이름 (COM 전용)"),
    backend: str = typer.Option("auto", "--backend", help="백엔드 선택 (auto/com/python-pptx)"),
    output_format: str = typer.Option("json", "--format-output", help="출력 형식 (json/text)"),
):
    """
    PowerPoint 슬라이드를 이미지로 내보냅니다.

    COM-First: Windows에서는 COM 백엔드 우선, python-pptx는 fallback

    **백엔드 선택**:
    - auto (기본): 자동으로 최적 백엔드 선택 (Windows COM 우선)
    - com: Windows COM 강제 사용 (완전한 기능)
    - python-pptx: python-pptx 강제 사용 (제한적 기능)

    **COM 백엔드 (Windows) - 완전한 기능!**:
    - ✅ Slide.Export() 사용
    - ✅ 다양한 이미지 형식 지원 (PNG, JPG, GIF, BMP, TIFF)
    - ✅ 해상도 조정 (DPI)
    - ✅ 크기 조정 (Width, Height)
    - ✅ 슬라이드 범위 선택

    **python-pptx 백엔드**:
    - ⚠️ 파일 저장 필수 (--file-path 필수)
    - Pillow를 사용한 이미지 생성 (제한적)

    **지원 이미지 형식**:
    - PNG: 고품질, 투명도 지원 (기본값)
    - JPG: 작은 파일 크기
    - GIF: 애니메이션 지원
    - BMP: 무손실
    - TIFF: 고품질 인쇄

    **파일명 패턴**:
    - {num}: 슬라이드 번호
    - {num:03d}: 3자리 숫자로 패딩 (예: 001, 002, ...)
    - {title}: 슬라이드 제목 (있는 경우)

    예제:
        # COM 백엔드 (활성 프레젠테이션 전체, PNG)
        oa ppt export-images --output-dir "slides"

        # 특정 슬라이드만 (1-10번)
        oa ppt export-images --output-dir "images" --slides "1-10"

        # 고해상도 JPG (300 DPI)
        oa ppt export-images --output-dir "export" --format JPG --dpi 300

        # 크기 지정 (1920x1080)
        oa ppt export-images --output-dir "hd" --width 1920 --height 1080

        # 커스텀 파일명
        oa ppt export-images --output-dir "out" --filename-pattern "page_{num:02d}" --presentation-name "report.pptx"
    """

    try:
        # 출력 디렉토리 검증
        normalized_output_dir = normalize_path(output_dir)
        output_path = Path(normalized_output_dir).resolve()
        output_path.mkdir(parents=True, exist_ok=True)

        # 이미지 형식 검증
        supported_formats = ["PNG", "JPG", "JPEG", "GIF", "BMP", "TIFF", "TIF"]
        image_format_upper = image_format.upper()

        if image_format_upper not in supported_formats:
            result = create_error_response(
                command="export-images",
                error=f"지원하지 않는 이미지 형식: {image_format}. 지원 형식: {', '.join(supported_formats)}",
                error_type="ValueError",
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

        # JPEG 정규화
        if image_format_upper == "JPEG":
            image_format_upper = "JPG"
        elif image_format_upper == "TIF":
            image_format_upper = "TIFF"

        # 백엔드 결정
        try:
            selected_backend = get_powerpoint_backend(force_backend=backend if backend != "auto" else None)
        except (ValueError, RuntimeError) as e:
            result = create_error_response(
                command="export-images",
                error=str(e),
                error_type=type(e).__name__,
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

        # 프레젠테이션 가져오기
        try:
            backend_inst, prs = get_or_open_presentation(
                file_path=file_path,
                presentation_name=presentation_name,
                backend=selected_backend,
            )
        except Exception as e:
            result = create_error_response(
                command="export-images",
                error=f"프레젠테이션을 열 수 없습니다: {str(e)}",
                error_type=type(e).__name__,
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

        # 백엔드별 처리
        if selected_backend == PowerPointBackend.COM.value:
            # COM 백엔드: 완전한 이미지 내보내기 기능
            try:
                total_slides = prs.Slides.Count

                # 슬라이드 범위 파싱
                if slides and slides.lower() != "all":
                    from .export_pdf import parse_slide_range

                    slide_numbers = parse_slide_range(slides, total_slides)

                    if not slide_numbers:
                        result = create_error_response(
                            command="export-images",
                            error=f"유효하지 않은 슬라이드 범위: {slides}",
                            error_type="ValueError",
                        )
                        typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
                        raise typer.Exit(1)
                else:
                    slide_numbers = list(range(1, total_slides + 1))

                # 이미지 내보내기
                exported_files = []
                total_size_bytes = 0

                for slide_num in slide_numbers:
                    slide = prs.Slides(slide_num)

                    # 파일명 생성
                    filename = filename_pattern.format(num=slide_num)
                    if not filename.lower().endswith(f".{image_format_upper.lower()}"):
                        filename += f".{image_format_upper.lower()}"

                    file_path_full = output_path / filename

                    # 슬라이드 내보내기
                    # Slide.Export(FileName, FilterName, ScaleWidth=0, ScaleHeight=0)
                    # ScaleWidth, ScaleHeight: 픽셀 단위 (0이면 기본 크기)

                    if width or height:
                        # 크기 지정
                        scale_width = width if width else 0
                        scale_height = height if height else 0
                        slide.Export(str(file_path_full), image_format_upper, scale_width, scale_height)
                    else:
                        # 기본 크기 (DPI 기반)
                        # PowerPoint 기본 슬라이드 크기: 10인치 × 7.5인치 (표준 4:3)
                        # 또는 13.333인치 × 7.5인치 (와이드 16:9)
                        slide_width_in = prs.PageSetup.SlideWidth / 72  # 포인트 → 인치
                        slide_height_in = prs.PageSetup.SlideHeight / 72

                        scale_width = int(slide_width_in * dpi)
                        scale_height = int(slide_height_in * dpi)

                        slide.Export(str(file_path_full), image_format_upper, scale_width, scale_height)

                    # 파일 크기 확인
                    file_size = file_path_full.stat().st_size
                    total_size_bytes += file_size

                    exported_files.append(
                        {
                            "slide_number": slide_num,
                            "filename": filename,
                            "file_path": str(file_path_full),
                            "file_size_kb": round(file_size / 1024, 2),
                        }
                    )

                total_size_mb = total_size_bytes / (1024 * 1024)

                # 성공 응답
                result_data = {
                    "backend": "com",
                    "output_directory": str(output_path),
                    "image_format": image_format_upper,
                    "total_slides": total_slides,
                    "exported_count": len(exported_files),
                    "total_size_mb": round(total_size_mb, 2),
                    "dpi": dpi,
                    "width": width,
                    "height": height,
                    "files": exported_files,
                }

                message = f"이미지 내보내기 완료 (COM): {len(exported_files)}개 슬라이드"

            except Exception as e:
                result = create_error_response(
                    command="export-images",
                    error=f"이미지 내보내기 실패: {str(e)}",
                    error_type=type(e).__name__,
                )
                typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
                raise typer.Exit(1)

        else:
            # python-pptx 백엔드
            if not file_path:
                result = create_error_response(
                    command="export-images",
                    error="python-pptx 백엔드는 --file-path 옵션이 필수입니다",
                    error_type="ValueError",
                )
                typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
                raise typer.Exit(1)

            result = create_error_response(
                command="export-images",
                error="python-pptx 백엔드는 이미지 내보내기를 직접 지원하지 않습니다. Pillow 등의 외부 라이브러리를 사용하거나, COM 백엔드를 사용하세요.",
                error_type="NotImplementedError",
                details={
                    "suggestions": [
                        "Use --backend com on Windows",
                        "Install Pillow for image manipulation",
                        "Use LibreOffice command line tools",
                    ]
                },
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

        # 성공 응답
        response = create_success_response(
            data=result_data,
            command="export-images",
            message=message,
        )

        # 출력
        if output_format == "json":
            typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
        else:
            typer.echo(f"✅ {message}")
            typer.echo(f"📁 디렉토리: {output_path}")
            typer.echo(f"🖼️ 형식: {image_format_upper}")
            typer.echo(f"📊 슬라이드: {result_data['exported_count']}개 / 총 {result_data['total_slides']}개")
            typer.echo(f"💾 총 크기: {result_data['total_size_mb']} MB")
            typer.echo(f"📐 해상도: {dpi} DPI")
            if width or height:
                typer.echo(f"📏 크기: {width or 'auto'} × {height or 'auto'} 픽셀")

    except typer.Exit:
        raise
    except Exception as e:
        result = create_error_response(
            command="export-images",
            error=str(e),
            error_type=type(e).__name__,
        )
        typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
        raise typer.Exit(1)
    finally:
        # python-pptx는 자동 정리, COM은 유지
        pass


if __name__ == "__main__":
    typer.run(export_images)
