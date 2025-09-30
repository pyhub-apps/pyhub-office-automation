"""
PowerPoint 새 프레젠테이션 생성 명령어 (Typer 버전)
AI 에이전트와의 연동을 위한 구조화된 출력 제공
"""

import json
import sys
from pathlib import Path
from typing import Optional

import typer

from pyhub_office_automation.version import get_version

from .utils import create_error_response, create_success_response, normalize_path


def presentation_create(
    save_path: Optional[str] = typer.Option(None, "--save-path", help="프레젠테이션을 저장할 경로"),
    template: str = typer.Option("blank", "--template", help="템플릿 선택 (blank, default)"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택"),
):
    """
    새로운 PowerPoint 프레젠테이션을 생성합니다.

    python-pptx를 사용하여 새 프레젠테이션을 생성하며,
    선택적으로 파일로 저장할 수 있습니다.

    예제:
        oa ppt presentation-create
        oa ppt presentation-create --save-path "report.pptx"
        oa ppt presentation-create --save-path "C:/Work/presentation.pptx" --template blank
    """
    try:
        # python-pptx import
        try:
            from pptx import Presentation
        except ImportError:
            result = create_error_response(
                command="presentation-create",
                error="python-pptx 패키지가 설치되지 않았습니다",
                error_type="ImportError",
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

        # 새 프레젠테이션 생성
        prs = Presentation()

        # 프레젠테이션 정보
        slide_count = len(prs.slides)
        layout_count = len(prs.slide_layouts)

        # 저장 경로가 지정된 경우 저장
        saved_path = None
        if save_path:
            try:
                # 경로 정규화
                save_path_obj = Path(normalize_path(save_path)).resolve()

                # 확장자가 없으면 .pptx 추가
                if not save_path_obj.suffix:
                    save_path_obj = save_path_obj.with_suffix(".pptx")

                # 디렉토리 생성 (필요한 경우)
                save_path_obj.parent.mkdir(parents=True, exist_ok=True)

                # 프레젠테이션 저장
                prs.save(str(save_path_obj))
                saved_path = str(save_path_obj)

            except Exception as e:
                result = create_error_response(
                    command="presentation-create",
                    error=f"프레젠테이션 저장 실패: {str(e)}",
                    error_type=type(e).__name__,
                )
                typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
                raise typer.Exit(1)

        # 성공 응답 데이터
        data = {
            "created": True,
            "template": template,
            "slide_count": slide_count,
            "layouts_available": layout_count,
            "saved": saved_path is not None,
        }

        if saved_path:
            data["file_path"] = saved_path
            data["file_size_bytes"] = Path(saved_path).stat().st_size

        result = create_success_response(
            command="presentation-create",
            data=data,
            message=f"프레젠테이션이 생성되었습니다" + (f": {saved_path}" if saved_path else ""),
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
