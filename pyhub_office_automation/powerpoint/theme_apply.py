"""
PowerPoint 테마 적용 명령어
프레젠테이션에 .thmx 테마 파일 또는 기본 테마를 적용합니다.
"""

import json
import platform
from pathlib import Path
from typing import Optional

import typer

from pyhub_office_automation.version import get_version

from .utils import create_error_response, create_success_response, normalize_path


def apply_theme_com(pptx_path: Path, theme_path: Path):
    """
    Windows COM을 사용하여 테마를 적용합니다.

    Args:
        pptx_path: 대상 프레젠테이션 경로
        theme_path: 테마 파일 경로 (.thmx)

    Returns:
        Dict: 적용 결과 정보

    Raises:
        ImportError: pywin32가 설치되지 않은 경우
        NotImplementedError: Windows가 아닌 경우
    """
    if platform.system() != "Windows":
        raise NotImplementedError("COM 인터페이스는 Windows에서만 사용 가능합니다")

    try:
        import win32com.client
    except ImportError:
        raise ImportError("pywin32 패키지가 필요합니다. 'pip install pywin32'로 설치하세요.")

    # PowerPoint Application 시작
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True

    try:
        # 프레젠테이션 열기
        presentation = powerpoint.Presentations.Open(str(pptx_path.resolve()), WithWindow=True)

        # 테마 적용
        presentation.ApplyTheme(str(theme_path.resolve()))

        # 저장
        presentation.Save()

        result = {
            "method": "COM",
            "theme_applied": True,
            "theme_file": str(theme_path),
        }

        # 프레젠테이션 닫기
        presentation.Close()

        return result

    except Exception as e:
        raise Exception(f"COM을 통한 테마 적용 실패: {str(e)}")

    finally:
        # PowerPoint 종료 (선택적)
        # powerpoint.Quit()
        pass


def apply_theme_python_pptx(pptx_path: Path):
    """
    python-pptx를 사용하여 제한적으로 테마를 적용합니다.

    Note: python-pptx는 .thmx 파일 적용을 직접 지원하지 않습니다.

    Args:
        pptx_path: 대상 프레젠테이션 경로

    Returns:
        Dict: 적용 결과 정보
    """
    try:
        from pptx import Presentation
    except ImportError:
        raise ImportError("python-pptx 패키지가 필요합니다. 'pip install python-pptx'로 설치하세요.")

    # 프레젠테이션 열기
    prs = Presentation(str(pptx_path))

    # python-pptx는 테마 직접 적용을 지원하지 않음
    result = {
        "method": "python-pptx",
        "theme_applied": False,
        "limitation": "python-pptx는 .thmx 테마 파일 적용을 지원하지 않습니다",
        "recommendation": "Windows 환경에서 pywin32를 설치하면 COM을 통해 테마를 적용할 수 있습니다",
    }

    return result


def theme_apply(
    file_path: str = typer.Option(..., "--file-path", help="PowerPoint 파일 경로"),
    theme_path: Optional[str] = typer.Option(None, "--theme-path", help="테마 파일 경로 (.thmx)"),
    theme_name: Optional[str] = typer.Option(None, "--theme-name", help="기본 테마 이름 (예: 'Office Theme')"),
    force_method: Optional[str] = typer.Option(None, "--method", help="강제 사용할 메서드 (com/python-pptx)"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 (json/text)"),
):
    """
    PowerPoint 프레젠테이션에 테마를 적용합니다.

    .thmx 파일을 통해 테마를 적용하거나, 기본 제공 테마를 적용할 수 있습니다.

    제약사항:
        - python-pptx는 테마 적용을 직접 지원하지 않습니다
        - Windows 환경에서 pywin32가 설치된 경우 COM을 통해 완전한 테마 적용이 가능합니다
        - macOS/Linux에서는 테마 적용이 제한됩니다

    예제:
        # Windows COM 사용 (완전한 기능)
        oa ppt theme-apply --file-path "report.pptx" --theme-path "corporate.thmx"

        # 기본 테마 적용 (Windows COM 전용)
        oa ppt theme-apply --file-path "report.pptx" --theme-name "Office Theme"

        # 크로스플랫폼 (제한적)
        oa ppt theme-apply --file-path "report.pptx" --method python-pptx
    """
    try:
        # 입력 검증
        if not theme_path and not theme_name:
            raise ValueError("--theme-path 또는 --theme-name 중 하나는 반드시 지정해야 합니다")

        if theme_path and theme_name:
            raise ValueError("--theme-path와 --theme-name은 동시에 사용할 수 없습니다")

        # 파일 경로 정규화 및 존재 확인
        normalized_pptx_path = normalize_path(file_path)
        pptx_path = Path(normalized_pptx_path).resolve()

        if not pptx_path.exists():
            raise FileNotFoundError(f"PowerPoint 파일을 찾을 수 없습니다: {pptx_path}")

        # 테마 파일 경로 확인
        theme_path_obj = None
        if theme_path:
            normalized_theme_path = normalize_path(theme_path)
            theme_path_obj = Path(normalized_theme_path).resolve()
            if not theme_path_obj.exists():
                raise FileNotFoundError(f"테마 파일을 찾을 수 없습니다: {theme_path_obj}")

        # 적용 방법 결정
        method = force_method
        if not method:
            # Windows + pywin32 설치되어 있으면 COM 사용
            if platform.system() == "Windows":
                try:
                    import win32com.client

                    method = "com"
                except ImportError:
                    method = "python-pptx"
            else:
                method = "python-pptx"

        # 테마 적용
        apply_result = {}

        if method == "com":
            if not theme_path_obj:
                raise ValueError("COM 메서드는 --theme-path가 필요합니다 (--theme-name은 현재 미지원)")

            apply_result = apply_theme_com(pptx_path, theme_path_obj)

        elif method == "python-pptx":
            apply_result = apply_theme_python_pptx(pptx_path)

        else:
            raise ValueError(f"지원하지 않는 메서드입니다: {method}")

        # 결과 데이터 구성
        result_data = {
            "file": str(pptx_path),
            "file_name": pptx_path.name,
            "theme_source": theme_path_obj.name if theme_path_obj else theme_name,
            "method": method,
            "platform": platform.system(),
            "apply_result": apply_result,
        }

        # 성공 응답
        if apply_result.get("theme_applied", False):
            message = f"테마 '{result_data['theme_source']}'를 적용했습니다"
        else:
            message = f"테마 적용이 제한되었습니다 (method: {method})"

        response = create_success_response(
            data=result_data,
            command="theme-apply",
            message=message,
        )

        # 출력
        if output_format == "json":
            typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
        else:
            if apply_result.get("theme_applied", False):
                typer.echo(f"✅ {message}")
                typer.echo(f"📄 파일: {pptx_path.name}")
                typer.echo(f"🎨 테마: {result_data['theme_source']}")
                typer.echo(f"⚙️  메서드: {method}")
            else:
                typer.echo(f"⚠️  {message}")
                typer.echo(f"📄 파일: {pptx_path.name}")
                typer.echo(f"⚙️  메서드: {method}")
                if apply_result.get("limitation"):
                    typer.echo(f"\n💡 제약사항: {apply_result['limitation']}")
                if apply_result.get("recommendation"):
                    typer.echo(f"💡 권장사항: {apply_result['recommendation']}")

    except FileNotFoundError as e:
        error_response = create_error_response(e, "theme-apply")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)

    except ValueError as e:
        error_response = create_error_response(e, "theme-apply")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)

    except (ImportError, NotImplementedError) as e:
        error_response = create_error_response(e, "theme-apply")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)

    except Exception as e:
        error_response = create_error_response(e, "theme-apply")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ 예기치 않은 오류: {str(e)}", err=True)
        raise typer.Exit(1)


if __name__ == "__main__":
    typer.run(theme_apply)
