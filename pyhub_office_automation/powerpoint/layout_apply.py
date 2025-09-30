"""
PowerPoint 레이아웃 적용 명령어 (COM-First)
슬라이드에 특정 레이아웃을 적용합니다.
"""

import json
from pathlib import Path
from typing import Optional, Union

import typer

from pyhub_office_automation.version import get_version

from .utils import (
    PowerPointBackend,
    create_error_response,
    create_success_response,
    get_layout_by_name_or_index,
    get_or_open_presentation,
    get_powerpoint_backend,
    normalize_path,
    validate_slide_number,
)


def layout_apply(
    slide_number: int = typer.Option(..., "--slide-number", help="레이아웃을 적용할 슬라이드 번호 (1부터 시작)"),
    layout: str = typer.Option(..., "--layout", help="레이아웃 이름 또는 인덱스 (예: 'Title Slide' 또는 0)"),
    file_path: Optional[str] = typer.Option(None, "--file-path", help="PowerPoint 파일 경로"),
    presentation_name: Optional[str] = typer.Option(None, "--presentation-name", help="열려있는 프레젠테이션 이름 (COM 전용)"),
    backend: str = typer.Option("auto", "--backend", help="백엔드 선택 (auto/com/python-pptx)"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 (json/text)"),
):
    """
    PowerPoint 슬라이드에 특정 레이아웃을 적용합니다.

    COM-First: Windows에서는 COM 백엔드 우선, python-pptx는 fallback

    **백엔드 선택**:
    - auto (기본): 자동으로 최적 백엔드 선택 (Windows COM 우선)
    - com: Windows COM 강제 사용 (완전한 기능 - 레이아웃 실제 변경 가능!)
    - python-pptx: python-pptx 강제 사용 (제한적 - 레이아웃 조회만 가능)

    **COM 백엔드 (Windows) - Issue #79 해결!**:
    - ✅ 슬라이드 레이아웃 실제 변경 가능!
    - Slide.CustomLayout 속성 사용
    - 열려있는 프레젠테이션에서 직접 작업

    **python-pptx 백엔드**:
    - ⚠️ 레이아웃 조회만 가능 (실제 변경 불가)
    - python-pptx API 제약사항

    예제:
        # COM 백엔드 (활성 프레젠테이션, 레이아웃 실제 변경)
        oa ppt layout-apply --slide-number 1 --layout "Title Slide"

        # COM 백엔드 (특정 프레젠테이션)
        oa ppt layout-apply --slide-number 2 --layout 1 --presentation-name "report.pptx"

        # python-pptx 백엔드 (레이아웃 조회만)
        oa ppt layout-apply --slide-number 1 --layout "Title Slide" --file-path "report.pptx" --backend python-pptx
    """
    backend_inst = None

    try:
        # 백엔드 결정
        try:
            selected_backend = get_powerpoint_backend(force_backend=backend if backend != "auto" else None)
        except (ValueError, RuntimeError) as e:
            result = create_error_response(
                command="layout-apply",
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
                command="layout-apply",
                error=f"프레젠테이션을 열 수 없습니다: {str(e)}",
                error_type=type(e).__name__,
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

        # 백엔드별 처리
        if selected_backend == PowerPointBackend.COM.value:
            # COM 백엔드: 레이아웃 실제 변경 가능!
            total_slides = prs.Slides.Count

            # 슬라이드 번호 검증 (COM은 1-based)
            if slide_number < 1 or slide_number > total_slides:
                result = create_error_response(
                    command="layout-apply",
                    error=f"슬라이드 번호가 범위를 벗어났습니다: {slide_number} (1-{total_slides})",
                    error_type="ValueError",
                )
                typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
                raise typer.Exit(1)

            slide = prs.Slides(slide_number)
            old_layout_name = slide.CustomLayout.Name

            # 레이아웃 찾기
            try:
                layout_index = int(layout)
                # python-pptx는 0-based, COM은 1-based
                if layout_index >= 0:
                    layout_index = layout_index + 1
            except ValueError:
                # 이름으로 찾기
                layout_found = False
                for i in range(1, prs.SlideMaster.CustomLayouts.Count + 1):
                    if prs.SlideMaster.CustomLayouts(i).Name == layout:
                        layout_index = i
                        layout_found = True
                        break
                if not layout_found:
                    result = create_error_response(
                        command="layout-apply",
                        error=f"레이아웃을 찾을 수 없습니다: {layout}",
                        error_type="ValueError",
                    )
                    typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
                    raise typer.Exit(1)

            # 레이아웃 실제 적용! (COM만 가능)
            try:
                new_layout = prs.SlideMaster.CustomLayouts(layout_index)
                slide.CustomLayout = new_layout
                new_layout_name = new_layout.Name

                # 성공 응답
                result_data = {
                    "backend": "com",
                    "slide_number": slide_number,
                    "old_layout": old_layout_name,
                    "new_layout": new_layout_name,
                    "applied": True,
                    "message": "레이아웃이 성공적으로 변경되었습니다!",
                }

                message = f"레이아웃 변경 완료 (COM): 슬라이드 {slide_number}, {old_layout_name} → {new_layout_name}"

            except Exception as e:
                result = create_error_response(
                    command="layout-apply",
                    error=f"레이아웃 적용 실패: {str(e)}",
                    error_type=type(e).__name__,
                )
                typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
                raise typer.Exit(1)

        else:
            # python-pptx 백엔드: 레이아웃 조회만 가능
            if not file_path:
                result = create_error_response(
                    command="layout-apply",
                    error="python-pptx 백엔드는 --file-path 옵션이 필수입니다",
                    error_type="ValueError",
                )
                typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
                raise typer.Exit(1)

            # 슬라이드 번호 검증
            slide_idx = validate_slide_number(slide_number, len(prs.slides))
            slide = prs.slides[slide_idx]

            # 기존 레이아웃 정보 저장
            old_layout_name = slide.slide_layout.name

            # 레이아웃 찾기
            try:
                layout_identifier = int(layout)
            except ValueError:
                layout_identifier = layout

            # 레이아웃 가져오기
            new_layout = get_layout_by_name_or_index(prs, layout_identifier)

            # python-pptx 제약: 레이아웃 변경 불가
            limitation_message = (
                "python-pptx는 기존 슬라이드의 레이아웃 직접 변경을 지원하지 않습니다. " "Windows에서 COM 백엔드를 사용하세요."
            )

            # 결과 데이터
            pptx_path = Path(normalize_path(file_path)).resolve()
            result_data = {
                "backend": "python-pptx",
                "file": str(pptx_path),
                "file_name": pptx_path.name,
                "slide_number": slide_number,
                "current_layout": old_layout_name,
                "requested_layout": new_layout.name,
                "layout_index": prs.slide_layouts.index(new_layout),
                "applied": False,
                "limitation": limitation_message,
                "alternative": "Windows에서 --backend com 사용",
            }

            message = f"레이아웃 조회 (python-pptx): 슬라이드 {slide_number}, 현재 {old_layout_name}"

        # 성공 응답
        response = create_success_response(
            data=result_data,
            command="layout-apply",
            message=message,
        )

        # 출력
        if output_format == "json":
            typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
        else:
            if result_data.get("applied"):
                typer.echo(f"✅ {message}")
                typer.echo(f"  슬라이드: {result_data['slide_number']}")
                typer.echo(f"  이전 레이아웃: {result_data['old_layout']}")
                typer.echo(f"  새 레이아웃: {result_data['new_layout']}")
            else:
                typer.echo(f"⚠️  {message}")
                typer.echo(f"  파일: {result_data.get('file_name', 'N/A')}")
                typer.echo(f"  슬라이드: {result_data['slide_number']}")
                typer.echo(f"  현재 레이아웃: {result_data.get('current_layout', 'N/A')}")
                typer.echo(f"  요청 레이아웃: {result_data.get('requested_layout', 'N/A')}")
                typer.echo(f"\n💡 제약사항: {result_data.get('limitation', 'N/A')}")
                typer.echo(f"💡 대안: {result_data.get('alternative', 'N/A')}")

    except typer.Exit:
        raise
    except Exception as e:
        result = create_error_response(
            command="layout-apply",
            error=str(e),
            error_type=type(e).__name__,
        )
        typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
        raise typer.Exit(1)
    finally:
        # python-pptx는 자동 정리, COM은 유지
        # COM 백엔드는 사용자가 명시적으로 닫아야 함
        pass


if __name__ == "__main__":
    typer.run(layout_apply)
