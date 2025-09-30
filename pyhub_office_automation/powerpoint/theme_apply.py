"""
PowerPoint 테마 적용 명령어 (COM-First)
프레젠테이션에 .thmx 테마 파일 또는 기본 테마를 적용합니다.
"""

import json
import platform
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


def theme_apply(
    theme_path: str = typer.Option(..., "--theme-path", help="테마 파일 경로 (.thmx)"),
    file_path: Optional[str] = typer.Option(None, "--file-path", help="PowerPoint 파일 경로"),
    presentation_name: Optional[str] = typer.Option(None, "--presentation-name", help="열려있는 프레젠테이션 이름 (COM 전용)"),
    backend: str = typer.Option("auto", "--backend", help="백엔드 선택 (auto/com/python-pptx)"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 (json/text)"),
):
    """
    PowerPoint 프레젠테이션에 테마를 적용합니다.

    COM-First: Windows에서는 COM 백엔드 우선, python-pptx는 fallback

    **백엔드 선택**:
    - auto (기본): 자동으로 최적 백엔드 선택 (Windows COM 우선)
    - com: Windows COM 강제 사용 (완전한 기능 - 테마 실제 적용 가능!)
    - python-pptx: python-pptx 강제 사용 (제한적 - 테마 적용 불가)

    **COM 백엔드 (Windows) - 완전한 기능!**:
    - ✅ .thmx 테마 파일 실제 적용 가능!
    - Presentation.ApplyTheme() 메서드 사용
    - 열려있는 프레젠테이션에서 직접 작업

    **python-pptx 백엔드**:
    - ⚠️ 테마 적용 불가 (API 제약사항)
    - COM 백엔드 사용 권장

    예제:
        # COM 백엔드 (활성 프레젠테이션, 테마 실제 적용)
        oa ppt theme-apply --theme-path "corporate.thmx"

        # COM 백엔드 (특정 프레젠테이션)
        oa ppt theme-apply --theme-path "corporate.thmx" --presentation-name "report.pptx"

        # python-pptx 백엔드 (테마 적용 불가, 제약사항 안내)
        oa ppt theme-apply --theme-path "corporate.thmx" --file-path "report.pptx" --backend python-pptx
    """
    backend_inst = None

    try:
        # 백엔드 결정
        try:
            selected_backend = get_powerpoint_backend(force_backend=backend if backend != "auto" else None)
        except (ValueError, RuntimeError) as e:
            result = create_error_response(
                command="theme-apply",
                error=str(e),
                error_type=type(e).__name__,
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

        # 테마 파일 경로 확인
        normalized_theme_path = normalize_path(theme_path)
        theme_path_obj = Path(normalized_theme_path).resolve()

        if not theme_path_obj.exists():
            result = create_error_response(
                command="theme-apply",
                error=f"테마 파일을 찾을 수 없습니다: {theme_path}",
                error_type="FileNotFoundError",
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
                command="theme-apply",
                error=f"프레젠테이션을 열 수 없습니다: {str(e)}",
                error_type=type(e).__name__,
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

        # 백엔드별 처리
        if selected_backend == PowerPointBackend.COM.value:
            # COM 백엔드: 테마 실제 적용 가능!
            try:
                # ApplyTheme 메서드로 테마 적용
                prs.ApplyTheme(str(theme_path_obj))

                # 성공 응답
                result_data = {
                    "backend": "com",
                    "theme_file": str(theme_path_obj),
                    "theme_name": theme_path_obj.name,
                    "applied": True,
                    "message": "테마가 성공적으로 적용되었습니다!",
                }

                message = f"테마 적용 완료 (COM): {theme_path_obj.name}"

            except Exception as e:
                result = create_error_response(
                    command="theme-apply",
                    error=f"테마 적용 실패: {str(e)}",
                    error_type=type(e).__name__,
                )
                typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
                raise typer.Exit(1)

        else:
            # python-pptx 백엔드: 테마 적용 불가
            limitation_message = (
                "python-pptx는 .thmx 테마 파일 적용을 지원하지 않습니다. " "Windows에서 COM 백엔드를 사용하세요."
            )

            result_data = {
                "backend": "python-pptx",
                "theme_file": str(theme_path_obj),
                "theme_name": theme_path_obj.name,
                "applied": False,
                "limitation": limitation_message,
                "alternative": "Windows에서 --backend com 사용",
            }

            message = f"테마 적용 불가 (python-pptx): {theme_path_obj.name}"

        # 성공 응답
        response = create_success_response(
            data=result_data,
            command="theme-apply",
            message=message,
        )

        # 출력
        if output_format == "json":
            typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
        else:
            if result_data.get("applied"):
                typer.echo(f"✅ {message}")
                typer.echo(f"  테마: {result_data['theme_name']}")
            else:
                typer.echo(f"⚠️  {message}")
                typer.echo(f"  테마: {result_data['theme_name']}")
                typer.echo(f"\n💡 제약사항: {result_data.get('limitation', 'N/A')}")
                typer.echo(f"💡 대안: {result_data.get('alternative', 'N/A')}")

    except typer.Exit:
        raise
    except Exception as e:
        result = create_error_response(
            command="theme-apply",
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
    typer.run(theme_apply)
