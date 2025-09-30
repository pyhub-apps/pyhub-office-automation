"""
PowerPoint 슬라이드쇼 시작 명령어 (COM 전용)
프로그래밍 방식으로 슬라이드쇼를 시작합니다.
"""

import json
import platform
from typing import Optional

import typer

from pyhub_office_automation.version import get_version

from .utils import PowerPointBackend, create_error_response, create_success_response, get_or_open_presentation


def slideshow_start(
    from_slide: int = typer.Option(1, "--from-slide", help="시작 슬라이드 번호 (1부터 시작, 기본: 1)"),
    end_slide: Optional[int] = typer.Option(None, "--end-slide", help="종료 슬라이드 번호 (기본: 마지막 슬라이드)"),
    show_type: str = typer.Option("speaker", "--show-type", help="쇼 타입 (speaker/window/kiosk, 기본: speaker)"),
    loop_until_stopped: bool = typer.Option(False, "--loop-until-stopped", help="ESC 누를 때까지 반복 (kiosk 모드)"),
    file_path: Optional[str] = typer.Option(None, "--file-path", help="PowerPoint 파일 경로"),
    presentation_name: Optional[str] = typer.Option(None, "--presentation-name", help="열려있는 프레젠테이션 이름 (COM 전용)"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 (json/text)"),
):
    """
    PowerPoint 슬라이드쇼를 시작합니다.

    Windows COM 전용 기능입니다. macOS/Linux에서는 지원되지 않습니다.

    **쇼 타입**:
    - speaker: 발표자 화면 (기본값, 전체화면)
    - window: 창 모드 (크기 조절 가능)
    - kiosk: 키오스크 모드 (자동 반복, ESC로 종료)

    **슬라이드 범위**:
    - --from-slide: 시작 슬라이드 (기본: 1)
    - --end-slide: 종료 슬라이드 (기본: 마지막 슬라이드)

    **반복 재생**:
    - --loop-until-stopped: ESC 누를 때까지 반복

    예제:
        # 활성 프레젠테이션의 첫 슬라이드부터 재생
        oa ppt slideshow-start

        # 특정 슬라이드부터 재생
        oa ppt slideshow-start --from-slide 3

        # 창 모드로 재생
        oa ppt slideshow-start --show-type window

        # 키오스크 모드 (자동 반복)
        oa ppt slideshow-start --show-type kiosk --loop-until-stopped

        # 슬라이드 범위 지정
        oa ppt slideshow-start --from-slide 2 --end-slide 5
    """

    # 1. 플랫폼 체크 (Windows 전용)
    if platform.system() != "Windows":
        result = create_error_response(
            command="slideshow-start",
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

    # 2. 쇼 타입 검증
    valid_show_types = ["speaker", "window", "kiosk"]
    if show_type not in valid_show_types:
        result = create_error_response(
            command="slideshow-start",
            error=f"올바르지 않은 쇼 타입: {show_type}. 유효한 값: {', '.join(valid_show_types)}",
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
            command="slideshow-start",
            error=f"프레젠테이션을 열 수 없습니다: {str(e)}",
            error_type=type(e).__name__,
        )
        typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
        raise typer.Exit(1)

    # 5. COM 슬라이드쇼 설정 및 시작
    try:
        # 슬라이드 총 개수 확인
        total_slides = prs.Slides.Count

        # 슬라이드 범위 검증
        if from_slide < 1 or from_slide > total_slides:
            result = create_error_response(
                command="slideshow-start",
                error=f"시작 슬라이드 번호가 범위를 벗어났습니다: {from_slide} (1-{total_slides})",
                error_type="ValueError",
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

        # 종료 슬라이드 기본값 처리 및 검증
        if end_slide is None:
            end_slide = total_slides
        elif end_slide < from_slide or end_slide > total_slides:
            result = create_error_response(
                command="slideshow-start",
                error=f"종료 슬라이드 번호가 잘못되었습니다: {end_slide} (범위: {from_slide}-{total_slides})",
                error_type="ValueError",
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

        # SlideShowSettings 가져오기
        settings = prs.SlideShowSettings

        # 슬라이드 범위 설정
        settings.RangeType = 3  # ppShowSlideRange = 3
        settings.StartingSlide = from_slide
        settings.EndingSlide = end_slide

        # 쇼 타입 설정
        # ppShowTypeSpeaker = 1 (발표자 화면)
        # ppShowTypeWindow = 2 (창 모드)
        # ppShowTypeKiosk = 3 (키오스크 모드)
        show_type_map = {
            "speaker": 1,
            "window": 2,
            "kiosk": 3,
        }
        settings.ShowType = show_type_map[show_type]

        # 반복 설정 (키오스크 모드에서 유용)
        if loop_until_stopped:
            settings.LoopUntilStopped = -1  # True
        else:
            settings.LoopUntilStopped = 0  # False

        # 슬라이드쇼 시작
        slideshow_window = settings.Run()

        # 결과 데이터
        result_data = {
            "backend": "com",
            "started": True,
            "from_slide": from_slide,
            "end_slide": end_slide,
            "total_slides": total_slides,
            "show_type": show_type,
            "loop_until_stopped": loop_until_stopped,
            "window_active": bool(slideshow_window),
        }

        message = f"슬라이드쇼 시작 완료 (슬라이드 {from_slide}-{end_slide}, {show_type} 모드)"

    except Exception as e:
        result = create_error_response(
            command="slideshow-start",
            error=f"슬라이드쇼 시작 실패: {str(e)}",
            error_type=type(e).__name__,
        )
        typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
        raise typer.Exit(1)

    # 6. 성공 응답
    response = create_success_response(
        data=result_data,
        command="slideshow-start",
        message=message,
    )

    # 출력
    if output_format == "json":
        typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
    else:
        typer.echo(f"✅ {message}")
        typer.echo(f"📊 총 슬라이드: {total_slides}")
        typer.echo(f"▶️ 재생 범위: {from_slide} - {end_slide}")
        typer.echo(f"🎭 쇼 타입: {show_type}")
        if loop_until_stopped:
            typer.echo("🔁 반복: 켜짐 (ESC로 종료)")


if __name__ == "__main__":
    typer.run(slideshow_start)
