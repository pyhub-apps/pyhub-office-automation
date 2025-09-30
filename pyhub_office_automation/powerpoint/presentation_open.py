"""
PowerPoint 프레젠테이션 열기 명령어 (COM-First)
기존 PowerPoint 파일을 열고 기본 정보 제공
"""

import json
from pathlib import Path

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


def presentation_open(
    file_path: str = typer.Option(..., "--file-path", help="열 프레젠테이션 파일의 경로"),
    backend: str = typer.Option("auto", "--backend", help="백엔드 선택 (auto/com/python-pptx)"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택"),
):
    """
    기존 PowerPoint 프레젠테이션 파일을 엽니다.

    COM-First: Windows에서는 COM 백엔드 우선, python-pptx는 fallback

    **백엔드 선택**:
    - auto (기본): 자동으로 최적 백엔드 선택 (Windows COM 우선)
    - com: Windows COM 강제 사용 (완전한 기능)
    - python-pptx: python-pptx 강제 사용 (제한적 기능)

    **COM 백엔드 (Windows)**:
    - 실제 PowerPoint 애플리케이션 실행
    - 모든 PowerPoint 기능 사용 가능
    - 열린 프레젠테이션 그대로 유지

    **python-pptx 백엔드**:
    - 파일을 메모리에 로드 (애플리케이션 실행 안 함)
    - 기본 정보만 조회 가능

    예제:
        oa ppt presentation-open --file-path "report.pptx"
        oa ppt presentation-open --file-path "report.pptx" --backend com
        oa ppt presentation-open --file-path "C:/Work/presentation.pptx"
    """
    backend_inst = None

    try:
        # 백엔드 결정
        try:
            selected_backend = get_powerpoint_backend(force_backend=backend if backend != "auto" else None)
        except (ValueError, RuntimeError) as e:
            result = create_error_response(
                command="presentation-open",
                error=str(e),
                error_type=type(e).__name__,
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

        # 경로 정규화 및 검증
        file_path_normalized = normalize_path(file_path)
        file_path_obj = Path(file_path_normalized).resolve()

        if not file_path_obj.exists():
            result = create_error_response(
                command="presentation-open",
                error=f"프레젠테이션 파일을 찾을 수 없습니다: {file_path}",
                error_type="FileNotFoundError",
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

        if not file_path_obj.is_file():
            result = create_error_response(
                command="presentation-open",
                error=f"경로가 파일이 아닙니다: {file_path}",
                error_type="ValueError",
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

        # 프레젠테이션 열기 (백엔드 자동 선택)
        try:
            backend_inst, prs = get_or_open_presentation(file_path=str(file_path_obj), backend=selected_backend)
        except Exception as e:
            result = create_error_response(
                command="presentation-open",
                error=f"프레젠테이션 파일을 열 수 없습니다: {str(e)}",
                error_type=type(e).__name__,
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

        # 프레젠테이션 정보 수집 (백엔드별 처리)
        file_stat = file_path_obj.stat()

        if selected_backend == PowerPointBackend.COM.value:
            # COM 백엔드: Presentation COM 객체
            data = {
                "backend": "com",
                "file_name": prs.Name,
                "file_path": prs.FullName,
                "file_size_bytes": file_stat.st_size,
                "slide_count": prs.Slides.Count,
                "saved": bool(prs.Saved),
                "read_only": bool(prs.ReadOnly) if hasattr(prs, "ReadOnly") else False,
            }

            # COM 백엔드는 프레젠테이션을 열어둔 상태로 유지
            message = f"프레젠테이션을 열었습니다 (COM): {prs.Name}"

            # 경고: PowerPoint 앱이 실행 중임을 알림
            if backend == "auto":
                message += " (PowerPoint 애플리케이션 실행 중)"

        else:
            # python-pptx 백엔드
            slide_count = len(prs.slides)
            layout_count = len(prs.slide_layouts)

            # 슬라이드 크기 (EMU -> inches 변환, 1 inch = 914400 EMU)
            slide_width_inches = prs.slide_width / 914400
            slide_height_inches = prs.slide_height / 914400

            data = {
                "backend": "python-pptx",
                "file_name": file_path_obj.name,
                "file_path": str(file_path_obj),
                "file_size_bytes": file_stat.st_size,
                "slide_count": slide_count,
                "layouts_available": layout_count,
                "slide_width_inches": round(slide_width_inches, 2),
                "slide_height_inches": round(slide_height_inches, 2),
            }

            message = f"프레젠테이션을 열었습니다 (python-pptx): {file_path_obj.name}"

            # python-pptx 제약사항 경고
            if backend == "auto":
                message += " ⚠️ 제한적 기능"

        # 성공 응답
        result = create_success_response(
            command="presentation-open",
            data=data,
            message=message,
        )

        typer.echo(json.dumps(result, ensure_ascii=False, indent=2))

    except typer.Exit:
        raise
    except Exception as e:
        result = create_error_response(
            command="presentation-open",
            error=str(e),
            error_type=type(e).__name__,
        )
        typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
        raise typer.Exit(1)
    finally:
        # python-pptx는 자동 정리, COM은 유지
        # COM 백엔드는 사용자가 명시적으로 닫아야 함
        pass
