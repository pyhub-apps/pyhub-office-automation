"""
PowerPoint VBA 매크로 실행 명령어 (COM 전용)
프레젠테이션에 포함된 VBA 매크로를 실행합니다.
"""

import json
import platform
from typing import List, Optional

import typer

from pyhub_office_automation.version import get_version

from .utils import PowerPointBackend, create_error_response, create_success_response, get_or_open_presentation


def run_macro(
    macro_name: str = typer.Option(..., "--macro-name", help="실행할 VBA 매크로 이름 (필수)"),
    args: Optional[str] = typer.Option(None, "--args", help="매크로 인자 (JSON 배열, 선택)"),
    file_path: Optional[str] = typer.Option(None, "--file-path", help="PowerPoint 파일 경로 (.pptm)"),
    presentation_name: Optional[str] = typer.Option(None, "--presentation-name", help="열려있는 프레젠테이션 이름 (COM 전용)"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 (json/text)"),
):
    """
    PowerPoint 프레젠테이션의 VBA 매크로를 실행합니다.

    Windows COM 전용 기능입니다. macOS/Linux에서는 지원되지 않습니다.

    **매크로 이름**:
    - 모듈명.프로시저명 형식 (예: "Module1.UpdateCharts")
    - 프로시저명만 사용 가능 (예: "FormatSlides")

    **매크로 인자**:
    - JSON 배열 형식으로 전달 (예: '["arg1", 123, true]')
    - 인자가 없는 매크로는 --args 생략

    **보안 경고**:
    - PowerPoint 매크로 보안 설정이 "모든 매크로 사용"으로 설정되어 있어야 합니다
    - 신뢰할 수 있는 문서만 실행하세요
    - .pptm 형식 파일 필요 (매크로 사용 프레젠테이션)

    예제:
        # 활성 프레젠테이션에서 매크로 실행
        oa ppt run-macro --macro-name "UpdateCharts"

        # 특정 파일에서 매크로 실행
        oa ppt run-macro --macro-name "Module1.FormatSlides" --file-path "report.pptm"

        # 인자가 있는 매크로 실행
        oa ppt run-macro --macro-name "ProcessData" --args '["Sheet1", 100, true]'

        # 열린 프레젠테이션에서 실행
        oa ppt run-macro --macro-name "ExportSlides" --presentation-name "report.pptm"
    """

    # 1. 플랫폼 체크 (Windows 전용)
    if platform.system() != "Windows":
        result = create_error_response(
            command="run-macro",
            error="이 명령어는 Windows에서만 사용 가능합니다 (COM 전용)",
            error_type="PlatformNotSupported",
            details={
                "platform": platform.system(),
                "alternative_suggestions": [
                    "Use PowerPoint application manually",
                    "Use Windows environment for full control",
                ],
            },
        )
        typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
        raise typer.Exit(1)

    # 2. 인자 파싱 (JSON 배열)
    macro_args = []
    if args:
        try:
            parsed_args = json.loads(args)
            if not isinstance(parsed_args, list):
                raise ValueError("매크로 인자는 JSON 배열 형식이어야 합니다")
            macro_args = parsed_args
        except json.JSONDecodeError as e:
            result = create_error_response(
                command="run-macro",
                error=f"매크로 인자 JSON 파싱 실패: {str(e)}",
                error_type="ValueError",
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

    # 3. 백엔드는 COM 고정
    selected_backend = PowerPointBackend.COM.value

    # 4. 프레젠테이션 가져오기
    try:
        backend_inst, prs = get_or_open_presentation(
            file_path=file_path,
            presentation_name=presentation_name,
            backend=selected_backend,
        )
    except Exception as e:
        result = create_error_response(
            command="run-macro",
            error=f"프레젠테이션을 열 수 없습니다: {str(e)}",
            error_type=type(e).__name__,
        )
        typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
        raise typer.Exit(1)

    # 5. COM을 통해 매크로 실행
    try:
        # PowerPoint Application 가져오기
        ppt_app = backend_inst.app

        # 매크로 이름 형식 확인 및 보정
        # "ModuleName.ProcedureName" 또는 "ProcedureName" 형식 지원
        # PowerPoint COM에서는 전체 경로 필요: PresentationName!ProcedureName
        prs_name = prs.Name

        # 확장자 제거 (.pptm → 없음)
        if "." in prs_name:
            prs_name_no_ext = prs_name.rsplit(".", 1)[0]
        else:
            prs_name_no_ext = prs_name

        # 매크로 전체 이름 구성
        if "." in macro_name:
            # 이미 모듈명.프로시저명 형식
            full_macro_name = f"{prs_name_no_ext}!{macro_name}"
        else:
            # 프로시저명만 있음
            full_macro_name = f"{prs_name_no_ext}!{macro_name}"

        # 매크로 실행
        # Application.Run(MacroName, Arg1, Arg2, ...)
        try:
            if macro_args:
                # 인자가 있는 경우
                result_value = ppt_app.Run(full_macro_name, *macro_args)
            else:
                # 인자가 없는 경우
                result_value = ppt_app.Run(full_macro_name)

            # 결과 데이터
            result_data = {
                "backend": "com",
                "macro_name": macro_name,
                "full_macro_name": full_macro_name,
                "args": macro_args,
                "executed": True,
                "result": str(result_value) if result_value is not None else None,
            }

            message = f"매크로 실행 완료: {macro_name}"

        except Exception as macro_error:
            # 매크로 실행 중 에러
            error_msg = str(macro_error)

            # 일반적인 에러 메시지 해석
            if "800A9C68" in error_msg or "can't find project or library" in error_msg.lower():
                hint = "매크로 보안 설정을 확인하세요. (파일 > 옵션 > 보안 센터 > 매크로 설정)"
            elif "800A9C64" in error_msg or "can't find macro" in error_msg.lower():
                hint = f"매크로를 찾을 수 없습니다: {macro_name}. 매크로 이름을 확인하세요."
            else:
                hint = "매크로 실행 중 에러가 발생했습니다. VBA 코드를 확인하세요."

            result = create_error_response(
                command="run-macro",
                error=f"매크로 실행 실패: {error_msg}",
                error_type=type(macro_error).__name__,
                details={
                    "macro_name": macro_name,
                    "full_macro_name": full_macro_name,
                    "hint": hint,
                    "security_guide": "PowerPoint 옵션 > 보안 센터 > 매크로 설정에서 '모든 매크로 사용' 또는 '디지털 서명한 매크로만 사용' 선택",
                },
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

    except Exception as e:
        result = create_error_response(
            command="run-macro",
            error=f"매크로 실행 중 예외 발생: {str(e)}",
            error_type=type(e).__name__,
        )
        typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
        raise typer.Exit(1)

    # 6. 성공 응답
    response = create_success_response(
        data=result_data,
        command="run-macro",
        message=message,
    )

    # 출력
    if output_format == "json":
        typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
    else:
        typer.echo(f"✅ {message}")
        typer.echo(f"📌 매크로: {full_macro_name}")
        if macro_args:
            typer.echo(f"📋 인자: {macro_args}")
        if result_data.get("result"):
            typer.echo(f"📤 결과: {result_data['result']}")


if __name__ == "__main__":
    typer.run(run_macro)
