"""
PowerPoint PDF 내보내기 명령어 (COM-First)
프레젠테이션을 PDF로 변환하여 저장합니다.
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


def export_pdf(
    output_path: str = typer.Option(..., "--output-path", help="PDF 저장 경로"),
    slides: Optional[str] = typer.Option(None, "--slides", help="내보낼 슬라이드 범위 (예: '1-5', '1,3,5', 'all')"),
    include_hidden: bool = typer.Option(False, "--include-hidden", help="숨겨진 슬라이드 포함"),
    embed_fonts: bool = typer.Option(True, "--embed-fonts/--no-embed-fonts", help="폰트 포함 (기본: True)"),
    quality: str = typer.Option("standard", "--quality", help="PDF 품질 (standard/print)"),
    file_path: Optional[str] = typer.Option(None, "--file-path", help="PowerPoint 파일 경로"),
    presentation_name: Optional[str] = typer.Option(None, "--presentation-name", help="열려있는 프레젠테이션 이름 (COM 전용)"),
    backend: str = typer.Option("auto", "--backend", help="백엔드 선택 (auto/com/python-pptx)"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 (json/text)"),
):
    """
    PowerPoint 프레젠테이션을 PDF로 내보냅니다.

    COM-First: Windows에서는 COM 백엔드 우선, python-pptx는 fallback

    **백엔드 선택**:
    - auto (기본): 자동으로 최적 백엔드 선택 (Windows COM 우선)
    - com: Windows COM 강제 사용 (완전한 기능)
    - python-pptx: python-pptx 강제 사용 (제한적 기능)

    **COM 백엔드 (Windows) - 완전한 기능!**:
    - ✅ Presentation.SaveAs() with ppSaveAsPDF format
    - ✅ 슬라이드 범위 선택 (전체, 특정 슬라이드, 범위)
    - ✅ 숨겨진 슬라이드 포함 여부 선택
    - ✅ 폰트 포함 옵션
    - ✅ 품질 설정 (Standard/Print)

    **python-pptx 백엔드**:
    - ⚠️ 파일 저장 필수 (--file-path 필수)
    - PDF 변환은 외부 도구 필요 (pypdf 또는 reportlab)
    - 제한적 기능

    **슬라이드 범위 지정**:
    - all 또는 생략: 전체 슬라이드
    - 1-5: 1번부터 5번 슬라이드
    - 1,3,5: 1번, 3번, 5번 슬라이드
    - 1-3,7,10-12: 복합 범위

    예제:
        # COM 백엔드 (활성 프레젠테이션 전체)
        oa ppt export-pdf --output-path "presentation.pdf"

        # 특정 슬라이드만 (1-10번)
        oa ppt export-pdf --output-path "slides_1-10.pdf" --slides "1-10"

        # 숨겨진 슬라이드 포함
        oa ppt export-pdf --output-path "full.pdf" --include-hidden

        # 고품질 인쇄용 PDF
        oa ppt export-pdf --output-path "print.pdf" --quality print --presentation-name "report.pptx"

        # python-pptx 백엔드
        oa ppt export-pdf --output-path "output.pdf" --file-path "input.pptx" --backend python-pptx
    """

    try:
        # 출력 경로 검증
        normalized_output_path = normalize_path(output_path)
        pdf_path = Path(normalized_output_path).resolve()

        # 디렉토리 생성
        pdf_path.parent.mkdir(parents=True, exist_ok=True)

        # 백엔드 결정
        try:
            selected_backend = get_powerpoint_backend(force_backend=backend if backend != "auto" else None)
        except (ValueError, RuntimeError) as e:
            result = create_error_response(
                command="export-pdf",
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
                command="export-pdf",
                error=f"프레젠테이션을 열 수 없습니다: {str(e)}",
                error_type=type(e).__name__,
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

        # 백엔드별 처리
        if selected_backend == PowerPointBackend.COM.value:
            # COM 백엔드: 완전한 PDF 내보내기 기능
            try:
                total_slides = prs.Slides.Count

                # 슬라이드 범위 파싱
                slide_range_type = "all"
                slide_numbers = []

                if slides and slides.lower() != "all":
                    slide_range_type = "custom"
                    slide_numbers = parse_slide_range(slides, total_slides)

                    if not slide_numbers:
                        result = create_error_response(
                            command="export-pdf",
                            error=f"유효하지 않은 슬라이드 범위: {slides}",
                            error_type="ValueError",
                        )
                        typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
                        raise typer.Exit(1)

                # ppSaveAsPDF = 32
                # ppPrintOutputSlides = 1 (기본)
                # ppPrintOutputBuildSlides = 2 (빌드 애니메이션 각각)
                # ppPrintOutputNotesPages = 5 (노트 페이지)

                # ppPrintHandoutVerticalFirst = 1
                # ppPrintHandoutHorizontalFirst = 2

                # ppPrintAll = 1
                # ppPrintSelection = 2 (현재 선택)
                # ppPrintCurrent = 3 (현재 슬라이드)
                # ppPrintSlideRange = 4 (범위 지정)

                # ppFixedFormatIntentScreen = 1 (Screen, 150 dpi)
                # ppFixedFormatIntentPrint = 2 (Print, 300 dpi)

                # 품질 설정
                if quality.lower() == "print":
                    intent = 2  # ppFixedFormatIntentPrint
                else:
                    intent = 1  # ppFixedFormatIntentScreen

                # 폰트 포함 설정
                embed_fonts_flag = -1 if embed_fonts else 0  # msoTrue / msoFalse

                # 숨겨진 슬라이드 포함 설정
                include_hidden_flag = -1 if include_hidden else 0

                # PDF 저장
                if slide_range_type == "all":
                    # 전체 슬라이드
                    prs.ExportAsFixedFormat(
                        Path=str(pdf_path),
                        FixedFormatType=2,  # ppFixedFormatTypePDF
                        Intent=intent,
                        FrameSlides=0,  # msoFalse
                        HandoutOrder=1,  # ppPrintHandoutVerticalFirst
                        OutputType=1,  # ppPrintOutputSlides
                        PrintHiddenSlides=include_hidden_flag,
                        PrintRange=None,
                        RangeType=1,  # ppPrintAll
                        SlideShowName="",
                        IncludeDocProperties=True,
                        KeepIRMSettings=True,
                        DocStructureTags=True,
                        BitmapMissingFonts=True,
                        UseISO19005_1=False,
                        ExternalExporter=None,
                    )
                else:
                    # 슬라이드 범위 지정
                    # PrintRange 객체 생성
                    print_range = prs.PrintOptions.Ranges.Add(
                        Start=min(slide_numbers),
                        End=max(slide_numbers),
                    )

                    prs.ExportAsFixedFormat(
                        Path=str(pdf_path),
                        FixedFormatType=2,  # ppFixedFormatTypePDF
                        Intent=intent,
                        FrameSlides=0,
                        HandoutOrder=1,
                        OutputType=1,
                        PrintHiddenSlides=include_hidden_flag,
                        PrintRange=print_range,
                        RangeType=4,  # ppPrintSlideRange
                        SlideShowName="",
                        IncludeDocProperties=True,
                        KeepIRMSettings=True,
                        DocStructureTags=True,
                        BitmapMissingFonts=True,
                        UseISO19005_1=False,
                        ExternalExporter=None,
                    )

                # PDF 파일 크기
                pdf_size_mb = pdf_path.stat().st_size / (1024 * 1024)

                # 성공 응답
                result_data = {
                    "backend": "com",
                    "output_file": str(pdf_path),
                    "output_file_name": pdf_path.name,
                    "file_size_mb": round(pdf_size_mb, 2),
                    "total_slides": total_slides,
                    "slide_range": slides or "all",
                    "slide_count": len(slide_numbers) if slide_numbers else total_slides,
                    "include_hidden": include_hidden,
                    "embed_fonts": embed_fonts,
                    "quality": quality,
                }

                message = f"PDF 내보내기 완료 (COM): {pdf_path.name}"
                if slide_range_type == "custom":
                    message += f", {len(slide_numbers)}개 슬라이드"

            except Exception as e:
                result = create_error_response(
                    command="export-pdf",
                    error=f"PDF 내보내기 실패: {str(e)}",
                    error_type=type(e).__name__,
                )
                typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
                raise typer.Exit(1)

        else:
            # python-pptx 백엔드
            if not file_path:
                result = create_error_response(
                    command="export-pdf",
                    error="python-pptx 백엔드는 --file-path 옵션이 필수입니다",
                    error_type="ValueError",
                )
                typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
                raise typer.Exit(1)

            result = create_error_response(
                command="export-pdf",
                error="python-pptx 백엔드는 PDF 내보내기를 직접 지원하지 않습니다. pypdf 또는 reportlab 등의 외부 라이브러리를 사용하거나, COM 백엔드를 사용하세요.",
                error_type="NotImplementedError",
                details={
                    "suggestions": [
                        "Use --backend com on Windows",
                        "Install pypdf or reportlab for PDF conversion",
                        "Export slides as images and convert to PDF",
                    ]
                },
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

        # 성공 응답
        response = create_success_response(
            data=result_data,
            command="export-pdf",
            message=message,
        )

        # 출력
        if output_format == "json":
            typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
        else:
            typer.echo(f"✅ {message}")
            typer.echo(f"📄 파일: {pdf_path}")
            typer.echo(f"💾 크기: {result_data['file_size_mb']} MB")
            typer.echo(f"📊 슬라이드: {result_data['slide_count']}개 / 총 {result_data['total_slides']}개")
            if include_hidden:
                typer.echo("👁️ 숨겨진 슬라이드: 포함")
            if embed_fonts:
                typer.echo("🔤 폰트: 포함")
            typer.echo(f"📐 품질: {quality.upper()}")

    except typer.Exit:
        raise
    except Exception as e:
        result = create_error_response(
            command="export-pdf",
            error=str(e),
            error_type=type(e).__name__,
        )
        typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
        raise typer.Exit(1)
    finally:
        # python-pptx는 자동 정리, COM은 유지
        pass


def parse_slide_range(range_str: str, total_slides: int) -> list[int]:
    """
    슬라이드 범위 문자열을 슬라이드 번호 리스트로 변환합니다.

    Args:
        range_str: 슬라이드 범위 (예: "1-5", "1,3,5", "1-3,7,10-12")
        total_slides: 전체 슬라이드 수

    Returns:
        슬라이드 번호 리스트 (1-based)
    """
    slide_numbers = set()

    # 쉼표로 분리
    parts = range_str.split(",")

    for part in parts:
        part = part.strip()

        if "-" in part:
            # 범위 (예: "1-5")
            try:
                start_str, end_str = part.split("-", 1)
                start = int(start_str.strip())
                end = int(end_str.strip())

                if start < 1 or end > total_slides or start > end:
                    return []

                slide_numbers.update(range(start, end + 1))
            except (ValueError, IndexError):
                return []
        else:
            # 단일 번호 (예: "3")
            try:
                num = int(part)
                if num < 1 or num > total_slides:
                    return []
                slide_numbers.add(num)
            except ValueError:
                return []

    return sorted(list(slide_numbers))


if __name__ == "__main__":
    typer.run(export_pdf)
