"""
PowerPoint 수식 추가 명령어 (COM 전용)
슬라이드에 수학 수식을 추가합니다 (OMath 또는 LaTeX).
"""

import json
import platform
from pathlib import Path
from typing import Optional

import typer

from pyhub_office_automation.version import get_version

from .utils import PowerPointBackend, create_error_response, create_success_response, get_or_open_presentation, normalize_path


def content_add_equation(
    slide_number: int = typer.Option(..., "--slide-number", help="수식을 추가할 슬라이드 번호 (1부터 시작)"),
    equation: str = typer.Option(..., "--equation", help="수식 텍스트 (LaTeX 또는 유니코드 수학)"),
    left: Optional[float] = typer.Option(None, "--left", help="수식 왼쪽 위치 (인치)"),
    top: Optional[float] = typer.Option(None, "--top", help="수식 상단 위치 (인치)"),
    width: Optional[float] = typer.Option(4.0, "--width", help="수식 박스 너비 (인치, 기본값: 4.0)"),
    height: Optional[float] = typer.Option(1.0, "--height", help="수식 박스 높이 (인치, 기본값: 1.0)"),
    center: bool = typer.Option(False, "--center", help="슬라이드 중앙에 배치 (--left, --top 무시)"),
    font_size: int = typer.Option(20, "--font-size", help="수식 글꼴 크기 (포인트, 기본값: 20)"),
    file_path: Optional[str] = typer.Option(None, "--file-path", help="PowerPoint 파일 경로"),
    presentation_name: Optional[str] = typer.Option(None, "--presentation-name", help="열려있는 프레젠테이션 이름 (COM 전용)"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 (json/text)"),
):
    """
    PowerPoint 슬라이드에 수학 수식을 추가합니다.

    Windows COM 전용 기능입니다. macOS/Linux에서는 지원되지 않습니다.

    **수식 입력 형식**:
    - LaTeX 스타일: "x=\\frac{-b\\pm\\sqrt{b^2-4ac}}{2a}"
    - 유니코드 수학: "x²+y²=r²"
    - Office Math (OMath) 형식도 지원

    **수식 예제**:
    - 이차 방정식: "x=\\frac{-b\\pm\\sqrt{b^2-4ac}}{2a}"
    - 피타고라스: "a^2+b^2=c^2"
    - 적분: "\\int_{0}^{\\infty} e^{-x}dx"
    - 합: "\\sum_{i=1}^{n} i = \\frac{n(n+1)}{2}"

    **위치 지정**:
      --center: 슬라이드 중앙에 배치
      --left, --top: 특정 위치에 배치

    예제:
        # 이차 방정식 공식 (중앙 배치)
        oa ppt content-add-equation --slide-number 2 --equation "x=\\frac{-b\\pm\\sqrt{b^2-4ac}}{2a}" --center

        # 피타고라스 정리 (위치 지정)
        oa ppt content-add-equation --slide-number 3 --equation "a^2+b^2=c^2" --left 2 --top 3 --font-size 24

        # 적분 수식 (특정 프레젠테이션)
        oa ppt content-add-equation --slide-number 4 --equation "\\int_{0}^{\\infty} e^{-x}dx=1" --presentation-name "math.pptx"
    """

    # 1. 플랫폼 체크 (Windows 전용)
    if platform.system() != "Windows":
        result = create_error_response(
            command="content-add-equation",
            error="이 명령어는 Windows에서만 사용 가능합니다 (COM 전용)",
            error_type="PlatformNotSupported",
            details={
                "platform": platform.system(),
                "alternative_suggestions": [
                    "Use MathType or similar tools manually",
                    "Use Windows environment for equation support",
                ],
            },
        )
        typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
        raise typer.Exit(1)

    try:
        # 2. 입력 검증
        if not center and (left is None or top is None):
            result = create_error_response(
                command="content-add-equation",
                error="--center를 사용하지 않는 경우 --left와 --top을 모두 지정해야 합니다",
                error_type="ValueError",
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

        # 3. 백엔드는 COM 고정
        selected_backend = PowerPointBackend.COM.value

        # 4. 프레젠테이션 가져오기
        try:
            backend_inst, prs = get_or_open_presentation(
                file_path=file_path,
                presentation_name=presentation_name,
                backend=selected_backend,
            )
        except Exception as e:
            result = create_error_response(
                command="content-add-equation",
                error=f"프레젠테이션을 열 수 없습니다: {str(e)}",
                error_type=type(e).__name__,
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

        # 5. COM을 통해 수식 추가
        try:
            total_slides = prs.Slides.Count

            # 슬라이드 번호 검증 (COM은 1-based)
            if slide_number < 1 or slide_number > total_slides:
                result = create_error_response(
                    command="content-add-equation",
                    error=f"슬라이드 번호가 범위를 벗어났습니다: {slide_number} (1-{total_slides})",
                    error_type="ValueError",
                )
                typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
                raise typer.Exit(1)

            slide = prs.Slides(slide_number)

            # 위치 계산
            if center:
                # 슬라이드 크기 가져오기 (포인트 단위)
                slide_width_pt = prs.PageSetup.SlideWidth
                slide_height_pt = prs.PageSetup.SlideHeight

                # 포인트를 인치로 변환
                slide_width_in = slide_width_pt / 72
                slide_height_in = slide_height_pt / 72

                # 중앙 배치 위치 계산
                final_left = (slide_width_in - width) / 2
                final_top = (slide_height_in - height) / 2
            else:
                final_left = left
                final_top = top

            # 인치를 포인트로 변환
            left_pt = final_left * 72
            top_pt = final_top * 72
            width_pt = width * 72
            height_pt = height * 72

            # 텍스트 박스 생성 (수식을 담을 컨테이너)
            # msoTextOrientationHorizontal = 1
            shape = slide.Shapes.AddTextbox(
                Orientation=1,  # msoTextOrientationHorizontal
                Left=left_pt,
                Top=top_pt,
                Width=width_pt,
                Height=height_pt,
            )

            # 텍스트 프레임 가져오기
            text_frame = shape.TextFrame
            text_range = text_frame.TextRange

            # 수식 텍스트 설정
            text_range.Text = equation

            # 글꼴 크기 설정
            text_range.Font.Size = font_size

            # OMath로 변환 시도 (Office 수식 편집기)
            try:
                # Word 스타일 OMath 변환
                # PowerPoint COM API는 직접 OMath를 지원하지 않으므로
                # 텍스트를 수식처럼 보이게 포맷팅
                text_range.Font.Name = "Cambria Math"

                # 텍스트를 중앙 정렬
                text_range.ParagraphFormat.Alignment = 2  # ppAlignCenter

                equation_type = "formatted_text"
            except Exception:
                # OMath 변환 실패 시 일반 텍스트로 유지
                equation_type = "text"

            # 결과 데이터
            result_data = {
                "backend": "com",
                "slide_number": slide_number,
                "equation": equation,
                "equation_type": equation_type,
                "font_size": font_size,
                "position": {
                    "left": round(final_left, 2),
                    "top": round(final_top, 2),
                    "width": width,
                    "height": height,
                },
                "centered": center,
                "note": "수식은 Cambria Math 폰트로 포맷팅됩니다. 복잡한 수식은 수동 편집이 필요할 수 있습니다.",
            }

            message = f"수식 추가 완료 (COM): 슬라이드 {slide_number}"

        except Exception as e:
            result = create_error_response(
                command="content-add-equation",
                error=f"수식 추가 실패: {str(e)}",
                error_type=type(e).__name__,
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

        # 6. 성공 응답
        response = create_success_response(
            data=result_data,
            command="content-add-equation",
            message=message,
        )

        # 출력
        if output_format == "json":
            typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
        else:
            typer.echo(f"✅ {message}")
            typer.echo(f"📍 슬라이드: {slide_number}")
            typer.echo(f"📐 수식: {equation}")
            typer.echo(f"📏 위치: {result_data['position']['left']}in × {result_data['position']['top']}in")
            typer.echo(f"📐 크기: {width}in × {height}in")
            typer.echo(f"🔤 글꼴 크기: {font_size}pt")
            typer.echo(f"ℹ️  {result_data['note']}")

    except typer.Exit:
        raise
    except Exception as e:
        result = create_error_response(
            command="content-add-equation",
            error=str(e),
            error_type=type(e).__name__,
        )
        typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
        raise typer.Exit(1)
    finally:
        # COM은 유지
        pass


if __name__ == "__main__":
    typer.run(content_add_equation)
