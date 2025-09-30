"""
PowerPoint 프레젠테이션 저장 명령어 (Typer 버전)
열려있는 프레젠테이션을 파일로 저장
"""

import json
import sys
from pathlib import Path

import typer

from pyhub_office_automation.version import get_version

from .utils import create_error_response, create_success_response, normalize_path


def presentation_save(
    source_path: str = typer.Option(..., "--source-path", help="저장할 원본 프레젠테이션 파일 경로"),
    save_path: str = typer.Option(..., "--save-path", help="프레젠테이션을 저장할 경로"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택"),
):
    """
    PowerPoint 프레젠테이션을 파일로 저장합니다.

    python-pptx 제한사항:
    - 열려있는 프레젠테이션을 추적할 수 없어 원본 파일 경로 필요
    - 실제로는 원본 파일을 열고 다른 경로로 저장하는 방식

    예제:
        oa ppt presentation-save --source-path "report.pptx" --save-path "report_backup.pptx"
        oa ppt presentation-save --source-path "C:/Work/original.pptx" --save-path "C:/Backup/copy.pptx"
    """
    try:
        # python-pptx import
        try:
            from pptx import Presentation
        except ImportError:
            result = create_error_response(
                command="presentation-save",
                error="python-pptx 패키지가 설치되지 않았습니다",
                error_type="ImportError",
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

        # 원본 경로 검증
        source_path_normalized = normalize_path(source_path)
        source_path_obj = Path(source_path_normalized).resolve()

        if not source_path_obj.exists():
            result = create_error_response(
                command="presentation-save",
                error=f"원본 프레젠테이션 파일을 찾을 수 없습니다: {source_path}",
                error_type="FileNotFoundError",
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

        # 저장 경로 정규화
        save_path_normalized = normalize_path(save_path)
        save_path_obj = Path(save_path_normalized).resolve()

        # 확장자가 없으면 .pptx 추가
        if not save_path_obj.suffix:
            save_path_obj = save_path_obj.with_suffix(".pptx")

        # 디렉토리 생성 (필요한 경우)
        save_path_obj.parent.mkdir(parents=True, exist_ok=True)

        # 프레젠테이션 열기
        try:
            prs = Presentation(str(source_path_obj))
        except Exception as e:
            result = create_error_response(
                command="presentation-save",
                error=f"원본 프레젠테이션을 열 수 없습니다: {str(e)}",
                error_type=type(e).__name__,
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

        # 프레젠테이션 저장
        try:
            prs.save(str(save_path_obj))
        except Exception as e:
            result = create_error_response(
                command="presentation-save",
                error=f"프레젠테이션 저장 실패: {str(e)}",
                error_type=type(e).__name__,
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

        # 저장된 파일 정보
        saved_file_stat = save_path_obj.stat()

        # 성공 응답 데이터
        data = {
            "source_path": str(source_path_obj),
            "saved_path": str(save_path_obj),
            "file_size_bytes": saved_file_stat.st_size,
            "slide_count": len(prs.slides),
        }

        result = create_success_response(
            command="presentation-save",
            data=data,
            message=f"프레젠테이션이 저장되었습니다: {save_path_obj.name}",
        )

        typer.echo(json.dumps(result, ensure_ascii=False, indent=2))

    except typer.Exit:
        raise
    except Exception as e:
        result = create_error_response(
            command="presentation-save",
            error=str(e),
            error_type=type(e).__name__,
        )
        typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
        raise typer.Exit(1)
