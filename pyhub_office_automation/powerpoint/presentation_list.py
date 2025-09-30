"""
PowerPoint 프레젠테이션 목록 조회 명령어 (Typer 버전)
열려있는 프레젠테이션 목록 조회 (Windows COM 전용)
"""

import json
import platform
import sys

import typer

from pyhub_office_automation.version import get_version

from .utils import create_error_response, create_success_response


def presentation_list(
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택"),
):
    """
    열려있는 PowerPoint 프레젠테이션 목록을 조회합니다.

    **제한사항**:
    - python-pptx는 열려있는 프레젠테이션 목록 조회를 지원하지 않습니다
    - Windows에서 pywin32 (COM)을 사용해야 가능합니다

    현재 구현:
    - Windows + pywin32: PowerPoint 애플리케이션의 열린 프레젠테이션 목록
    - 기타 플랫폼: 기능 제한 안내 메시지

    예제:
        oa ppt presentation-list
        oa ppt presentation-list --format json
    """
    current_platform = platform.system()

    # Windows에서 COM 사용 시도
    if current_platform == "Windows":
        try:
            import win32com.client

            try:
                # PowerPoint 애플리케이션에 연결
                powerpoint = win32com.client.Dispatch("PowerPoint.Application")

                # 열린 프레젠테이션 목록
                presentations_data = []

                for pres in powerpoint.Presentations:
                    try:
                        pres_info = {
                            "name": pres.Name,
                            "full_name": pres.FullName,
                            "slide_count": pres.Slides.Count,
                            "saved": pres.Saved,
                        }
                        presentations_data.append(pres_info)
                    except Exception as e:
                        # 개별 프레젠테이션 정보 수집 실패
                        presentations_data.append(
                            {
                                "name": getattr(pres, "Name", "Unknown"),
                                "error": f"정보 수집 실패: {str(e)}",
                            }
                        )

                # 성공 응답
                data = {
                    "platform": "Windows (COM)",
                    "presentation_count": len(presentations_data),
                    "presentations": presentations_data,
                }

                result = create_success_response(
                    command="presentation-list",
                    data=data,
                    message=(
                        f"{len(presentations_data)}개의 프레젠테이션이 열려있습니다"
                        if presentations_data
                        else "열려있는 프레젠테이션이 없습니다"
                    ),
                )

                typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
                return

            except Exception as e:
                # PowerPoint 애플리케이션에 연결 실패
                result = create_error_response(
                    command="presentation-list",
                    error=f"PowerPoint 애플리케이션에 연결할 수 없습니다: {str(e)}",
                    error_type=type(e).__name__,
                )
                typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
                raise typer.Exit(1)

        except ImportError:
            # pywin32가 설치되지 않음
            result = create_error_response(
                command="presentation-list",
                error="Windows COM 기능을 사용하려면 pywin32 패키지가 필요합니다. 'pip install pywin32'로 설치하세요",
                error_type="ImportError",
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

    # 비-Windows 플랫폼
    else:
        data = {
            "platform": current_platform,
            "presentation_count": 0,
            "presentations": [],
            "limitation": "python-pptx는 열려있는 프레젠테이션 목록 조회를 지원하지 않습니다",
            "alternative": "파일 경로를 직접 지정하여 presentation-info 명령어를 사용하세요",
        }

        result = create_success_response(
            command="presentation-list",
            data=data,
            message=f"{current_platform}에서는 presentation-list 기능이 제한됩니다. Windows COM 전용 기능입니다.",
        )

        typer.echo(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    typer.run(presentation_list)
