"""
PowerPoint 프레젠테이션 열기 명령어 (Typer 버전)
기존 PowerPoint 파일을 열고 기본 정보 제공
"""

import json
import sys
from pathlib import Path

import typer

from pyhub_office_automation.version import get_version

from .utils import create_error_response, create_success_response, normalize_path


def presentation_open(
    file_path: str = typer.Option(..., "--file-path", help="열 프레젠테이션 파일의 경로"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택"),
):
    """
    기존 PowerPoint 프레젠테이션 파일을 엽니다.

    python-pptx를 사용하여 파일을 열고 기본 정보를 반환합니다.
    실제로는 파일을 메모리에 로드하는 작업으로, PowerPoint 애플리케이션을 실행하지 않습니다.

    예제:
        oa ppt presentation-open --file-path "report.pptx"
        oa ppt presentation-open --file-path "C:/Work/presentation.pptx"
    """
    try:
        # python-pptx import
        try:
            from pptx import Presentation
        except ImportError:
            result = create_error_response(
                command="presentation-open",
                error="python-pptx 패키지가 설치되지 않았습니다",
                error_type="ImportError",
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

        # 프레젠테이션 열기
        try:
            prs = Presentation(str(file_path_obj))
        except Exception as e:
            result = create_error_response(
                command="presentation-open",
                error=f"프레젠테이션 파일을 열 수 없습니다: {str(e)}",
                error_type=type(e).__name__,
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

        # 프레젠테이션 정보 수집
        slide_count = len(prs.slides)
        layout_count = len(prs.slide_layouts)

        # 슬라이드 크기 (EMU -> inches 변환, 1 inch = 914400 EMU)
        slide_width_inches = prs.slide_width / 914400
        slide_height_inches = prs.slide_height / 914400

        # 파일 정보
        file_stat = file_path_obj.stat()

        # 성공 응답 데이터
        data = {
            "file_name": file_path_obj.name,
            "file_path": str(file_path_obj),
            "file_size_bytes": file_stat.st_size,
            "slide_count": slide_count,
            "layouts_available": layout_count,
            "slide_width_inches": round(slide_width_inches, 2),
            "slide_height_inches": round(slide_height_inches, 2),
        }

        result = create_success_response(
            command="presentation-open",
            data=data,
            message=f"프레젠테이션을 열었습니다: {file_path_obj.name}",
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
