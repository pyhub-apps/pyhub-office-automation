"""
PowerPoint 프레젠테이션 저장 명령어 (COM-First)
열려있는 프레젠테이션을 파일로 저장
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
    create_success_response,
    get_or_open_presentation,
    get_powerpoint_backend,
    normalize_path,
)


def presentation_save(
    save_path: str = typer.Option(..., "--save-path", help="프레젠테이션을 저장할 경로"),
    source_path: Optional[str] = typer.Option(None, "--source-path", help="원본 프레젠테이션 파일 경로 (python-pptx 전용)"),
    presentation_name: Optional[str] = typer.Option(
        None, "--presentation-name", help="저장할 열려있는 프레젠테이션 이름 (COM 전용)"
    ),
    backend: str = typer.Option("auto", "--backend", help="백엔드 선택 (auto/com/python-pptx)"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택"),
):
    """
    PowerPoint 프레젠테이션을 파일로 저장합니다.

    COM-First: Windows에서는 COM 백엔드 우선, python-pptx는 fallback

    **백엔드 선택**:
    - auto (기본): 자동으로 최적 백엔드 선택 (Windows COM 우선)
    - com: Windows COM 강제 사용 (완전한 기능)
    - python-pptx: python-pptx 강제 사용 (제한적 기능)

    **COM 백엔드 (Windows)**:
    - 현재 열려있는 프레젠테이션을 저장 (--presentation-name 또는 활성 프레젠테이션)
    - 이미 열려있는 PowerPoint 파일에서 직접 저장 (빠르고 효율적)
    - SaveAs 메서드 사용

    **python-pptx 백엔드**:
    - --source-path 필수 (열려있는 프레젠테이션 추적 불가)
    - 원본 파일을 열고 다른 경로로 저장하는 방식

    예제:
        # COM 백엔드 (활성 프레젠테이션 저장)
        oa ppt presentation-save --save-path "report_backup.pptx"

        # COM 백엔드 (특정 프레젠테이션 저장)
        oa ppt presentation-save --save-path "backup.pptx" --presentation-name "report.pptx"

        # python-pptx 백엔드
        oa ppt presentation-save --save-path "backup.pptx" --source-path "report.pptx" --backend python-pptx
    """
    backend_inst = None

    try:
        # 백엔드 결정
        try:
            selected_backend = get_powerpoint_backend(force_backend=backend if backend != "auto" else None)
        except (ValueError, RuntimeError) as e:
            result = create_error_response(
                command="presentation-save",
                error=str(e),
                error_type=type(e).__name__,
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

        # 백엔드별 처리
        if selected_backend == PowerPointBackend.COM.value:
            # COM 백엔드: 열려있는 프레젠테이션 저장
            try:
                # 프레젠테이션 가져오기 (이름 또는 활성)
                backend_inst, prs = get_or_open_presentation(
                    file_path=source_path if source_path else None,
                    presentation_name=presentation_name,
                    backend=selected_backend,
                )

                # SaveAs로 저장
                prs.SaveAs(str(save_path_obj))

                # 저장된 파일 정보
                data = {
                    "backend": "com",
                    "source_name": prs.Name,
                    "source_path": prs.FullName if prs.Path else None,
                    "saved_path": str(save_path_obj),
                    "file_size_bytes": save_path_obj.stat().st_size if save_path_obj.exists() else None,
                    "slide_count": prs.Slides.Count,
                }

                message = f"프레젠테이션이 저장되었습니다 (COM): {save_path_obj.name}"

            except Exception as e:
                result = create_error_response(
                    command="presentation-save",
                    error=f"프레젠테이션 저장 실패: {str(e)}",
                    error_type=type(e).__name__,
                )
                typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
                raise typer.Exit(1)

        else:
            # python-pptx 백엔드: source_path 필수
            if not source_path:
                result = create_error_response(
                    command="presentation-save",
                    error="python-pptx 백엔드는 --source-path 옵션이 필수입니다",
                    error_type="ValueError",
                )
                typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
                raise typer.Exit(1)

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

            # 프레젠테이션 열기 및 저장
            try:
                prs = Presentation(str(source_path_obj))
                prs.save(str(save_path_obj))

                # 저장된 파일 정보
                data = {
                    "backend": "python-pptx",
                    "source_path": str(source_path_obj),
                    "saved_path": str(save_path_obj),
                    "file_size_bytes": save_path_obj.stat().st_size,
                    "slide_count": len(prs.slides),
                }

                message = f"프레젠테이션이 저장되었습니다 (python-pptx): {save_path_obj.name}"

            except Exception as e:
                result = create_error_response(
                    command="presentation-save",
                    error=f"프레젠테이션 저장 실패: {str(e)}",
                    error_type=type(e).__name__,
                )
                typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
                raise typer.Exit(1)

        # 성공 응답
        result = create_success_response(
            command="presentation-save",
            data=data,
            message=message,
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
    finally:
        # python-pptx는 자동 정리, COM은 유지
        # COM 백엔드는 사용자가 명시적으로 닫아야 함
        pass
