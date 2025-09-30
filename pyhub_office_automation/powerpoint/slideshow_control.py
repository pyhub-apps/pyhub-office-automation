"""
PowerPoint 슬라이드쇼 제어 명령어 (COM 전용)
실행 중인 슬라이드쇼를 프로그래밍 방식으로 제어합니다.
"""

import json
import platform
from typing import Optional

import typer

from pyhub_office_automation.version import get_version

from .utils import create_error_response, create_success_response


def slideshow_control(
    action: str = typer.Option(..., "--action", help="제어 액션 (next/previous/goto/end, 필수)"),
    slide: Optional[int] = typer.Option(None, "--slide", help="goto 액션 시 이동할 슬라이드 번호"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 (json/text)"),
):
    """
    실행 중인 PowerPoint 슬라이드쇼를 제어합니다.

    Windows COM 전용 기능입니다. macOS/Linux에서는 지원되지 않습니다.

    **액션**:
    - next: 다음 슬라이드로 이동
    - previous: 이전 슬라이드로 이동
    - goto: 특정 슬라이드로 이동 (--slide 옵션 필수)
    - end: 슬라이드쇼 종료

    **주의사항**:
    - 슬라이드쇼가 실행 중이어야 합니다
    - goto 액션 사용 시 --slide 옵션 필수

    예제:
        # 다음 슬라이드로 이동
        oa ppt slideshow-control --action next

        # 이전 슬라이드로 이동
        oa ppt slideshow-control --action previous

        # 5번 슬라이드로 이동
        oa ppt slideshow-control --action goto --slide 5

        # 슬라이드쇼 종료
        oa ppt slideshow-control --action end
    """

    # 1. 플랫폼 체크 (Windows 전용)
    if platform.system() != "Windows":
        result = create_error_response(
            command="slideshow-control",
            error="이 명령어는 Windows에서만 사용 가능합니다 (COM 전용)",
            error_type="PlatformNotSupported",
            details={
                "platform": platform.system(),
                "alternative_suggestions": [
                    "Use PowerPoint application manually",
                    "Use Windows environment for full control",
                ],
            },
        )
        typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
        raise typer.Exit(1)

    # 2. 액션 검증
    valid_actions = ["next", "previous", "goto", "end"]
    if action not in valid_actions:
        result = create_error_response(
            command="slideshow-control",
            error=f"올바르지 않은 액션: {action}. 유효한 값: {', '.join(valid_actions)}",
            error_type="ValueError",
        )
        typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
        raise typer.Exit(1)

    # 3. goto 액션 시 slide 옵션 필수 검증
    if action == "goto" and slide is None:
        result = create_error_response(
            command="slideshow-control",
            error="goto 액션을 사용하려면 --slide 옵션이 필요합니다",
            error_type="ValueError",
        )
        typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
        raise typer.Exit(1)

    # 4. COM 초기화 및 슬라이드쇼 윈도우 찾기
    try:
        import pythoncom
        import win32com.client

        # PowerPoint Application 가져오기
        try:
            ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
        except Exception:
            result = create_error_response(
                command="slideshow-control",
                error="PowerPoint가 실행 중이지 않습니다",
                error_type="RuntimeError",
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

        # 슬라이드쇼 윈도우 확인
        if ppt_app.SlideShowWindows.Count == 0:
            result = create_error_response(
                command="slideshow-control",
                error="실행 중인 슬라이드쇼가 없습니다. slideshow-start 명령으로 먼저 슬라이드쇼를 시작하세요.",
                error_type="RuntimeError",
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

        # 첫 번째 슬라이드쇼 윈도우 가져오기
        slideshow_window = ppt_app.SlideShowWindows(1)
        view = slideshow_window.View

    except ImportError:
        result = create_error_response(
            command="slideshow-control",
            error="pywin32 패키지가 설치되지 않았습니다. 'pip install pywin32'로 설치하세요",
            error_type="ImportError",
        )
        typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
        raise typer.Exit(1)

    # 5. 액션 실행
    try:
        current_slide_before = view.Slide.SlideIndex

        if action == "next":
            view.Next()
            message = "다음 슬라이드로 이동"
        elif action == "previous":
            view.Previous()
            message = "이전 슬라이드로 이동"
        elif action == "goto":
            # 슬라이드 범위 검증
            total_slides = slideshow_window.Presentation.Slides.Count
            if slide < 1 or slide > total_slides:
                result = create_error_response(
                    command="slideshow-control",
                    error=f"슬라이드 번호가 범위를 벗어났습니다: {slide} (1-{total_slides})",
                    error_type="ValueError",
                )
                typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
                raise typer.Exit(1)

            view.GotoSlide(slide)
            message = f"{slide}번 슬라이드로 이동"
        elif action == "end":
            view.Exit()
            message = "슬라이드쇼 종료"

            # 종료 시에는 current_slide_after를 None으로 설정
            result_data = {
                "backend": "com",
                "action": action,
                "slideshow_ended": True,
            }

            response = create_success_response(
                data=result_data,
                command="slideshow-control",
                message=message,
            )

            if output_format == "json":
                typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
            else:
                typer.echo(f"✅ {message}")
            return

        # end 액션이 아닌 경우 현재 슬라이드 정보 가져오기
        try:
            current_slide_after = view.Slide.SlideIndex
        except Exception:
            # 슬라이드쇼가 종료된 경우
            current_slide_after = None

        result_data = {
            "backend": "com",
            "action": action,
            "slide_before": current_slide_before,
            "slide_after": current_slide_after,
        }

    except Exception as e:
        result = create_error_response(
            command="slideshow-control",
            error=f"슬라이드쇼 제어 실패: {str(e)}",
            error_type=type(e).__name__,
        )
        typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
        raise typer.Exit(1)

    # 6. 성공 응답
    response = create_success_response(
        data=result_data,
        command="slideshow-control",
        message=message,
    )

    # 출력
    if output_format == "json":
        typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
    else:
        typer.echo(f"✅ {message}")
        if action != "end":
            typer.echo(f"📊 슬라이드: {current_slide_before} → {current_slide_after}")


if __name__ == "__main__":
    typer.run(slideshow_control)
