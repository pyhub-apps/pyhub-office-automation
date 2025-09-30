"""
PowerPoint 애니메이션 효과 추가 명령어 (COM 전용)
슬라이드 객체에 애니메이션 효과를 추가합니다.
"""

import json
import platform
from typing import Optional

import typer

from pyhub_office_automation.version import get_version

from .utils import PowerPointBackend, create_error_response, create_success_response, get_or_open_presentation

# 애니메이션 효과 상수 매핑
# msoAnimation* 상수들
ANIMATION_EFFECTS = {
    # 입장 효과 (Entrance)
    "appear": 1,  # msoAnimEffectAppear
    "fade": 10,  # msoAnimEffectFade
    "fly-in": 2,  # msoAnimEffectFlyIn
    "float-in": 3,  # msoAnimEffectFloatIn
    "split": 4,  # msoAnimEffectSplit
    "wipe": 5,  # msoAnimEffectWipe
    "shape": 6,  # msoAnimEffectShape
    "wheel": 7,  # msoAnimEffectWheel
    "random-bars": 8,  # msoAnimEffectRandomBars
    "grow-turn": 9,  # msoAnimEffectGrowAndTurn
    "zoom": 11,  # msoAnimEffectZoom
    "swivel": 12,  # msoAnimEffectSwivel
    "bounce": 13,  # msoAnimEffectBounce
    # 강조 효과 (Emphasis)
    "pulse": 15,  # msoAnimEffectPulse
    "color-pulse": 16,  # msoAnimEffectColorPulse
    "teeter": 17,  # msoAnimEffectTeeter
    "spin": 18,  # msoAnimEffectSpin
    "grow-shrink": 19,  # msoAnimEffectGrowShrink
    "desaturate": 20,  # msoAnimEffectDesaturate
    "lighten": 21,  # msoAnimEffectLighten
    "transparency": 22,  # msoAnimEffectTransparency
    # 종료 효과 (Exit)
    "fly-out": 23,  # msoAnimEffectFlyOut
    "float-out": 24,  # msoAnimEffectFloatOut
    "disappear": 25,  # msoAnimEffectDisappear
}

# 애니메이션 방향 상수
ANIMATION_DIRECTIONS = {
    "from-bottom": 1,  # msoAnimDirectionFromBottom
    "from-top": 2,  # msoAnimDirectionFromTop
    "from-left": 3,  # msoAnimDirectionFromLeft
    "from-right": 4,  # msoAnimDirectionFromRight
    "from-bottom-left": 5,  # msoAnimDirectionFromBottomLeft
    "from-bottom-right": 6,  # msoAnimDirectionFromBottomRight
    "from-top-left": 7,  # msoAnimDirectionFromTopLeft
    "from-top-right": 8,  # msoAnimDirectionFromTopRight
}

# 애니메이션 트리거 타입
ANIMATION_TRIGGERS = {
    "on-click": 1,  # msoAnimTriggerOnPageClick
    "with-previous": 2,  # msoAnimTriggerWithPrevious
    "after-previous": 3,  # msoAnimTriggerAfterPrevious
}


def animation_add(
    slide_number: int = typer.Option(..., "--slide-number", help="슬라이드 번호 (1부터 시작, 필수)"),
    shape_index: int = typer.Option(..., "--shape-index", help="도형 인덱스 (1부터 시작, 필수)"),
    effect: str = typer.Option(..., "--effect", help="애니메이션 효과 이름 (예: fade, fly-in, zoom)"),
    duration: float = typer.Option(1.0, "--duration", help="지속 시간(초, 기본: 1.0)"),
    trigger: str = typer.Option(
        "on-click", "--trigger", help="트리거 타입 (on-click/with-previous/after-previous, 기본: on-click)"
    ),
    direction: Optional[str] = typer.Option(None, "--direction", help="방향 (from-bottom/from-top/from-left/from-right 등)"),
    file_path: Optional[str] = typer.Option(None, "--file-path", help="PowerPoint 파일 경로"),
    presentation_name: Optional[str] = typer.Option(None, "--presentation-name", help="열려있는 프레젠테이션 이름 (COM 전용)"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 (json/text)"),
):
    """
    PowerPoint 슬라이드의 객체에 애니메이션 효과를 추가합니다.

    Windows COM 전용 기능입니다. macOS/Linux에서는 지원되지 않습니다.

    **지원 효과 (일부)**:

    입장 효과 (Entrance):
    - appear, fade, fly-in, float-in, split, wipe, shape, wheel
    - random-bars, grow-turn, zoom, swivel, bounce

    강조 효과 (Emphasis):
    - pulse, color-pulse, teeter, spin, grow-shrink
    - desaturate, lighten, transparency

    종료 효과 (Exit):
    - fly-out, float-out, disappear

    **방향 (direction)**:
    - from-bottom, from-top, from-left, from-right
    - from-bottom-left, from-bottom-right, from-top-left, from-top-right

    **트리거 (trigger)**:
    - on-click: 클릭 시 재생 (기본값)
    - with-previous: 이전 효과와 동시 재생
    - after-previous: 이전 효과 종료 후 자동 재생

    예제:
        # 기본 Fade 효과
        oa ppt animation-add --slide-number 1 --shape-index 1 --effect fade

        # FlyIn 효과 (아래에서 위로, 2초 지속)
        oa ppt animation-add --slide-number 2 --shape-index 1 --effect fly-in --direction from-bottom --duration 2.0

        # Zoom 효과 (이전 효과 후 자동 재생)
        oa ppt animation-add --slide-number 3 --shape-index 2 --effect zoom --trigger after-previous --duration 1.5

        # Spin 강조 효과
        oa ppt animation-add --slide-number 4 --shape-index 1 --effect spin --duration 1.0
    """

    # 1. 플랫폼 체크 (Windows 전용)
    if platform.system() != "Windows":
        result = create_error_response(
            command="animation-add",
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

    # 2. 효과 이름 검증
    effect_lower = effect.lower()
    if effect_lower not in ANIMATION_EFFECTS:
        result = create_error_response(
            command="animation-add",
            error=f"지원하지 않는 효과: {effect}",
            error_type="ValueError",
            details={
                "supported_effects": list(ANIMATION_EFFECTS.keys()),
                "hint": "oa ppt animation-add --help 를 실행하여 지원 효과 목록을 확인하세요",
            },
        )
        typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
        raise typer.Exit(1)

    # 3. 트리거 검증
    trigger_lower = trigger.lower()
    if trigger_lower not in ANIMATION_TRIGGERS:
        result = create_error_response(
            command="animation-add",
            error=f"올바르지 않은 트리거: {trigger}. 유효한 값: {', '.join(ANIMATION_TRIGGERS.keys())}",
            error_type="ValueError",
        )
        typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
        raise typer.Exit(1)

    # 4. 방향 검증 (선택적)
    direction_value = None
    if direction:
        direction_lower = direction.lower()
        if direction_lower not in ANIMATION_DIRECTIONS:
            result = create_error_response(
                command="animation-add",
                error=f"올바르지 않은 방향: {direction}. 유효한 값: {', '.join(ANIMATION_DIRECTIONS.keys())}",
                error_type="ValueError",
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)
        direction_value = ANIMATION_DIRECTIONS[direction_lower]

    # 5. 백엔드는 COM 고정
    selected_backend = PowerPointBackend.COM.value

    # 6. 프레젠테이션 가져오기
    try:
        backend_inst, prs = get_or_open_presentation(
            file_path=file_path,
            presentation_name=presentation_name,
            backend=selected_backend,
        )
    except Exception as e:
        result = create_error_response(
            command="animation-add",
            error=f"프레젠테이션을 열 수 없습니다: {str(e)}",
            error_type=type(e).__name__,
        )
        typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
        raise typer.Exit(1)

    # 7. COM을 통해 애니메이션 추가
    try:
        # 슬라이드 가져오기
        total_slides = prs.Slides.Count
        if slide_number < 1 or slide_number > total_slides:
            result = create_error_response(
                command="animation-add",
                error=f"슬라이드 번호가 범위를 벗어났습니다: {slide_number} (1-{total_slides})",
                error_type="ValueError",
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

        slide = prs.Slides(slide_number)

        # 도형 가져오기
        total_shapes = slide.Shapes.Count
        if shape_index < 1 or shape_index > total_shapes:
            result = create_error_response(
                command="animation-add",
                error=f"도형 인덱스가 범위를 벗어났습니다: {shape_index} (1-{total_shapes})",
                error_type="ValueError",
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

        shape = slide.Shapes(shape_index)

        # 타임라인 가져오기
        timeline = slide.TimeLine
        main_sequence = timeline.MainSequence

        # 애니메이션 효과 추가
        effect_value = ANIMATION_EFFECTS[effect_lower]
        anim_effect = main_sequence.AddEffect(Shape=shape, effectId=effect_value, trigger=ANIMATION_TRIGGERS[trigger_lower])

        # 지속 시간 설정
        anim_effect.Timing.Duration = duration

        # 방향 설정 (지원하는 효과만)
        if direction_value is not None:
            try:
                anim_effect.EffectParameters.Direction = direction_value
            except Exception:
                # 일부 효과는 방향을 지원하지 않음 (무시)
                pass

        # 결과 데이터
        result_data = {
            "backend": "com",
            "slide_number": slide_number,
            "shape_index": shape_index,
            "effect": effect_lower,
            "effect_code": effect_value,
            "duration": duration,
            "trigger": trigger_lower,
            "direction": direction_lower if direction else None,
            "animation_count": main_sequence.Count,
        }

        message = f"애니메이션 효과 추가 완료: {effect} (슬라이드 {slide_number}, 도형 {shape_index})"

    except Exception as e:
        result = create_error_response(
            command="animation-add",
            error=f"애니메이션 추가 실패: {str(e)}",
            error_type=type(e).__name__,
        )
        typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
        raise typer.Exit(1)

    # 8. 성공 응답
    response = create_success_response(
        data=result_data,
        command="animation-add",
        message=message,
    )

    # 출력
    if output_format == "json":
        typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
    else:
        typer.echo(f"✅ {message}")
        typer.echo(f"📍 슬라이드: {slide_number}")
        typer.echo(f"🎭 도형: {shape_index}")
        typer.echo(f"✨ 효과: {effect_lower}")
        typer.echo(f"⏱️ 지속시간: {duration}초")
        typer.echo(f"🎬 트리거: {trigger_lower}")
        if direction:
            typer.echo(f"➡️ 방향: {direction_lower}")
        typer.echo(f"📊 총 애니메이션: {result_data['animation_count']}개")


if __name__ == "__main__":
    typer.run(animation_add)
