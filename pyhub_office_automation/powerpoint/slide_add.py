"""
PowerPoint 슬라이드 추가 명령어 (COM-First)
새 슬라이드를 지정된 레이아웃으로 추가
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
    get_layout_by_name_or_index,
    get_or_open_presentation,
    get_powerpoint_backend,
    normalize_path,
    validate_slide_number,
)


def slide_add(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="프레젠테이션 파일 경로"),
    presentation_name: Optional[str] = typer.Option(None, "--presentation-name", help="열려있는 프레젠테이션 이름 (COM 전용)"),
    layout: Optional[str] = typer.Option(None, "--layout", help="레이아웃 이름 또는 인덱스 (기본: 1=Title and Content)"),
    position: Optional[int] = typer.Option(None, "--position", help="삽입 위치 (1-based, 기본: 끝에 추가)"),
    backend: str = typer.Option("auto", "--backend", help="백엔드 선택 (auto/com/python-pptx)"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택 (json/text)"),
):
    """
    PowerPoint 프레젠테이션에 새 슬라이드를 추가합니다.

    COM-First: Windows에서는 COM 백엔드 우선, python-pptx는 fallback

    **백엔드 선택**:
    - auto (기본): 자동으로 최적 백엔드 선택 (Windows COM 우선)
    - com: Windows COM 강제 사용 (완전한 기능)
    - python-pptx: python-pptx 강제 사용 (제한적 기능)

    **COM 백엔드 (Windows)**:
    - 열려있는 프레젠테이션에 슬라이드 추가 (--presentation-name 또는 --file-path)
    - Slides.Add() 메서드 사용
    - 위치 지정 가능

    **python-pptx 백엔드**:
    - --file-path 필수
    - slides.add_slide() 메서드 사용
    - 위치 지정 시 XML 조작 필요

    예제:
        # COM 백엔드 (활성 프레젠테이션)
        oa ppt slide-add

        # COM 백엔드 (특정 프레젠테이션)
        oa ppt slide-add --presentation-name "report.pptx" --layout "Blank"

        # python-pptx 백엔드
        oa ppt slide-add --file-path "report.pptx" --layout "Blank" --position 3 --backend python-pptx
    """
    backend_inst = None

    try:
        # 백엔드 결정
        try:
            selected_backend = get_powerpoint_backend(force_backend=backend if backend != "auto" else None)
        except (ValueError, RuntimeError) as e:
            result = create_error_response(
                command="slide-add",
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
                command="slide-add",
                error=f"프레젠테이션을 열 수 없습니다: {str(e)}",
                error_type=type(e).__name__,
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

        # 백엔드별 처리
        if selected_backend == PowerPointBackend.COM.value:
            # COM 백엔드: Slides.Add() 사용
            total_slides = prs.Slides.Count

            # 레이아웃 결정 (기본값: 1 = Title and Content)
            if layout is None:
                layout_index = 2  # COM은 1-based, 2는 Title and Content
            else:
                # 숫자 문자열이면 int로 변환
                try:
                    layout_index = int(layout)
                    # python-pptx는 0-based, COM은 1-based이므로 +1
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
                            command="slide-add",
                            error=f"레이아웃을 찾을 수 없습니다: {layout}",
                            error_type="ValueError",
                        )
                        typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
                        raise typer.Exit(1)

            # 위치 결정 (COM은 1-based)
            if position is not None:
                insert_position = position
            else:
                insert_position = total_slides + 1  # 끝에 추가

            try:
                # COM Slides.Add(Index, Layout)
                new_slide = prs.Slides.Add(insert_position, layout_index)
                layout_name = new_slide.CustomLayout.Name
                final_position = new_slide.SlideIndex

                # 성공 응답
                data = {
                    "backend": "com",
                    "slide_number": final_position,
                    "layout": layout_name,
                    "total_slides": prs.Slides.Count,
                }

                message = f"슬라이드가 추가되었습니다 (COM): 위치 {final_position}, 레이아웃 {layout_name}"

            except Exception as e:
                result = create_error_response(
                    command="slide-add",
                    error=f"슬라이드 추가 실패: {str(e)}",
                    error_type=type(e).__name__,
                )
                typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
                raise typer.Exit(1)

        else:
            # python-pptx 백엔드
            if not file_path:
                result = create_error_response(
                    command="slide-add",
                    error="python-pptx 백엔드는 --file-path 옵션이 필수입니다",
                    error_type="ValueError",
                )
                typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
                raise typer.Exit(1)

            total_slides = len(prs.slides)

            # 레이아웃 결정 (기본값: 1 = Title and Content)
            if layout is None:
                layout_identifier = 1  # 기본 레이아웃 인덱스
            else:
                # 숫자 문자열이면 int로 변환
                try:
                    layout_identifier = int(layout)
                except ValueError:
                    layout_identifier = layout

            # 레이아웃 찾기
            try:
                slide_layout = get_layout_by_name_or_index(prs, layout_identifier)
                layout_name = slide_layout.name
            except (ValueError, IndexError) as e:
                result = create_error_response(
                    command="slide-add",
                    error=str(e),
                    error_type=type(e).__name__,
                )
                typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
                raise typer.Exit(1)

            # 위치 검증 (끝에 추가 허용)
            if position is not None:
                target_idx = validate_slide_number(position, total_slides, allow_append=True)
            else:
                target_idx = None  # 끝에 추가

            # 슬라이드 추가 (끝에)
            new_slide = prs.slides.add_slide(slide_layout)
            new_slide_idx = len(prs.slides) - 1  # 0-based index

            # 위치 지정이 있고, 끝이 아니면 순서 조정
            if position is not None and target_idx < new_slide_idx:
                # XML 레벨에서 순서 조정
                xml_slides = prs.slides._sldIdLst
                slide_id_element = xml_slides[new_slide_idx]
                xml_slides.remove(slide_id_element)
                xml_slides.insert(target_idx, slide_id_element)
                final_position = position
            else:
                final_position = len(prs.slides)

            # 파일 저장
            file_path_obj = Path(normalize_path(file_path)).resolve()
            prs.save(str(file_path_obj))

            # 성공 응답
            data = {
                "backend": "python-pptx",
                "slide_number": final_position,
                "layout": layout_name,
                "total_slides": len(prs.slides),
            }

            message = f"슬라이드가 추가되었습니다 (python-pptx): 위치 {final_position}, 레이아웃 {layout_name}"

        result = create_success_response(
            command="slide-add",
            data=data,
            message=message,
        )

        if output_format == "json":
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
        else:
            typer.echo(f"슬라이드 추가 완료")
            typer.echo(f"  위치: {data['slide_number']}번")
            typer.echo(f"  레이아웃: {data['layout']}")
            typer.echo(f"  총 슬라이드: {data['total_slides']}개")

    except typer.Exit:
        raise
    except Exception as e:
        result = create_error_response(
            command="slide-add",
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
    typer.run(slide_add)
