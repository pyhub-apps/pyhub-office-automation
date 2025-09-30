"""
PowerPoint 새 프레젠테이션 생성 명령어 (COM-First)
새 PowerPoint 파일 생성 및 저장
"""

import json
import sys
from pathlib import Path
from typing import Optional

import typer

from pyhub_office_automation.version import get_version

from .utils import (
    PowerPointBackend,
    create_error_response,
    create_presentation_with_backend,
    create_success_response,
    get_powerpoint_backend,
    normalize_path,
)


def presentation_create(
    save_path: Optional[str] = typer.Option(None, "--save-path", help="프레젠테이션을 저장할 경로"),
    template: str = typer.Option("blank", "--template", help="템플릿 선택 (blank, default)"),
    backend: str = typer.Option("auto", "--backend", help="백엔드 선택 (auto/com/python-pptx)"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택"),
):
    """
    새로운 PowerPoint 프레젠테이션을 생성합니다.

    COM-First: Windows에서는 COM 백엔드 우선, python-pptx는 fallback

    **백엔드 선택**:
    - auto (기본): 자동으로 최적 백엔드 선택 (Windows COM 우선)
    - com: Windows COM 강제 사용 (완전한 기능)
    - python-pptx: python-pptx 강제 사용 (제한적 기능)

    **COM 백엔드 (Windows)**:
    - 실제 PowerPoint 애플리케이션 실행
    - 모든 PowerPoint 기능 사용 가능
    - 생성된 프레젠테이션 그대로 유지 (저장 시에만 파일로 저장)

    **python-pptx 백엔드**:
    - 파일을 메모리에 생성 (애플리케이션 실행 안 함)
    - 기본 프레젠테이션 생성 가능

    예제:
        oa ppt presentation-create
        oa ppt presentation-create --save-path "report.pptx"
        oa ppt presentation-create --save-path "C:/Work/presentation.pptx" --backend com
    """
    backend_inst = None

    try:
        # 백엔드 결정
        try:
            selected_backend = get_powerpoint_backend(force_backend=backend if backend != "auto" else None)
        except (ValueError, RuntimeError) as e:
            result = create_error_response(
                command="presentation-create",
                error=str(e),
                error_type=type(e).__name__,
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

        # 템플릿 검증
        if template not in ["blank", "default"]:
            result = create_error_response(
                command="presentation-create",
                error=f"지원하지 않는 템플릿입니다: {template}. 사용 가능: blank, default",
                error_type="ValueError",
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

        # 저장 경로 처리
        save_path_obj = None
        if save_path:
            # 경로 정규화
            save_path_obj = Path(normalize_path(save_path)).resolve()

            # 확장자가 없으면 .pptx 추가
            if not save_path_obj.suffix:
                save_path_obj = save_path_obj.with_suffix(".pptx")

            # 디렉토리 생성 (필요한 경우)
            save_path_obj.parent.mkdir(parents=True, exist_ok=True)

        # 프레젠테이션 생성 (백엔드별 처리)
        try:
            backend_inst, prs = create_presentation_with_backend(
                save_path=str(save_path_obj) if save_path_obj else None, backend=selected_backend
            )
        except Exception as e:
            result = create_error_response(
                command="presentation-create",
                error=f"프레젠테이션 생성 실패: {str(e)}",
                error_type=type(e).__name__,
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

        # 프레젠테이션 정보 수집 (백엔드별 처리)
        if selected_backend == PowerPointBackend.COM.value:
            # COM 백엔드: Presentation COM 객체
            data = {
                "backend": "com",
                "created": True,
                "template": template,
                "file_name": prs.Name,
                "slide_count": prs.Slides.Count,
                "saved": bool(prs.Saved),
            }

            if save_path_obj:
                data["file_path"] = str(save_path_obj)
                if save_path_obj.exists():
                    data["file_size_bytes"] = save_path_obj.stat().st_size

            # COM 백엔드는 프레젠테이션을 열어둔 상태로 유지
            message = f"프레젠테이션이 생성되었습니다 (COM): {prs.Name}"
            if save_path_obj:
                message += f" - 저장됨: {save_path_obj}"

            # 경고: PowerPoint 앱이 실행 중임을 알림
            if backend == "auto":
                message += " (PowerPoint 애플리케이션 실행 중)"

        else:
            # python-pptx 백엔드
            slide_count = len(prs.slides)
            layout_count = len(prs.slide_layouts)

            data = {
                "backend": "python-pptx",
                "created": True,
                "template": template,
                "slide_count": slide_count,
                "layouts_available": layout_count,
                "saved": save_path_obj is not None,
            }

            if save_path_obj:
                data["file_path"] = str(save_path_obj)
                if save_path_obj.exists():
                    data["file_size_bytes"] = save_path_obj.stat().st_size

            message = f"프레젠테이션이 생성되었습니다 (python-pptx)"
            if save_path_obj:
                message += f": {save_path_obj}"

            # python-pptx 제약사항 경고
            if backend == "auto":
                message += " ⚠️ 제한적 기능"

        # 성공 응답
        result = create_success_response(
            command="presentation-create",
            data=data,
            message=message,
        )

        typer.echo(json.dumps(result, ensure_ascii=False, indent=2))

    except typer.Exit:
        raise
    except Exception as e:
        result = create_error_response(
            command="presentation-create",
            error=str(e),
            error_type=type(e).__name__,
        )
        typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
        raise typer.Exit(1)
    finally:
        # python-pptx는 자동 정리, COM은 유지
        # COM 백엔드는 사용자가 명시적으로 닫아야 함
        pass
