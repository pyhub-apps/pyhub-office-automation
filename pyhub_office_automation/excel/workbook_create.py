"""
Excel 새 워크북 생성 명령어 (Engine 기반)
"""

import json
import sys
from pathlib import Path
from typing import Optional

import typer

from pyhub_office_automation.version import get_version

from .engines import get_engine
from .utils import ExecutionTimer, create_error_response, create_success_response, normalize_path


def workbook_create(
    name: str = typer.Option("NewWorkbook", "--name", help="생성할 워크북의 이름 (참고용)"),
    save_path: Optional[str] = typer.Option(None, "--save-path", help="워크북을 저장할 경로"),
    visible: bool = typer.Option(True, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택"),
):
    """
    새로운 Excel 워크북을 생성합니다.

    현재 실행 중인 Excel 애플리케이션에 새 워크북을 추가합니다.

    예제:
        oa excel workbook-create
        oa excel workbook-create --save-path "data.xlsx"
        oa excel workbook-create --save-path "C:/Reports/monthly.xlsx"
    """
    try:
        # 실행 시간 측정 시작
        with ExecutionTimer() as timer:
            # Engine 획득
            engine = get_engine()

            # 저장 경로 처리
            save_path_str = None
            if save_path:
                # 경로 정규화
                save_path_obj = Path(normalize_path(save_path)).resolve()

                # 확장자가 없으면 .xlsx 추가
                if not save_path_obj.suffix:
                    save_path_obj = save_path_obj.with_suffix(".xlsx")

                # 디렉토리 생성 (필요한 경우)
                save_path_obj.parent.mkdir(parents=True, exist_ok=True)

                save_path_str = str(save_path_obj)

            # Engine을 통해 새 워크북 생성
            book = engine.create_workbook(save_path=save_path_str, visible=visible)

            # 워크북 정보 가져오기
            wb_info = engine.get_workbook_info(book)

            # 시트 목록 구성
            sheets_info = []
            for idx, sheet_name in enumerate(wb_info["sheets"]):
                sheets_info.append({"name": sheet_name, "index": idx, "is_active": sheet_name == wb_info["active_sheet"]})

            # 워크북 정보 구성
            workbook_info = {
                "name": normalize_path(wb_info["name"]),
                "full_name": normalize_path(wb_info["full_name"]),
                "saved": wb_info["saved"],
                "saved_path": save_path_str,
                "sheet_count": wb_info["sheet_count"],
                "active_sheet": wb_info["active_sheet"],
                "sheets": sheets_info,
            }

            # 데이터 구성
            data_content = {
                "workbook": workbook_info,
                "creation_method": "engine",
            }

            # 성공 메시지
            if save_path_str:
                message = f"새 워크북 '{workbook_info['name']}'을(를) 생성하고 '{save_path_str}'에 저장했습니다"
            else:
                message = f"새 워크북 '{workbook_info['name']}'을(를) 생성했습니다"

            # 성공 응답 생성
            response = create_success_response(
                data=data_content,
                command="workbook-create",
                message=message,
                execution_time_ms=timer.execution_time_ms,
            )

            # 출력 형식에 따른 결과 반환
            if output_format == "json":
                typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
            else:  # text 형식
                wb = workbook_info
                typer.echo(f"✅ {message}")
                typer.echo()
                typer.echo(f"📁 워크북명: {wb['name']}")
                typer.echo(f"📍 전체경로: {wb['full_name']}")
                if save_path_str:
                    typer.echo(f"💾 저장경로: {save_path_str}")
                    typer.echo(f"💾 저장상태: {'저장됨' if wb['saved'] else '저장되지 않음'}")
                else:
                    typer.echo(f"⚠️  저장되지 않은 새 워크북 (필요시 직접 저장하세요)")

                typer.echo(f"📄 시트 수: {wb['sheet_count']}")
                typer.echo(f"📑 활성 시트: {wb['active_sheet']}")

                typer.echo()
                typer.echo("📋 생성된 시트:")
                for i, sheet in enumerate(wb["sheets"], 1):
                    active_mark = " (활성)" if sheet.get("is_active") else ""
                    typer.echo(f"  {i}. {sheet['name']}{active_mark}")

                if not save_path_str:
                    typer.echo()
                    typer.echo("💡 워크북을 저장하려면 Excel에서 Ctrl+S를 누르거나")
                    typer.echo("   --save-path 옵션으로 경로를 지정하세요")

    except Exception as e:
        error_response = create_error_response(e, "workbook-create")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
            typer.echo("💡 Excel이 실행되고 있는지 확인하세요.", err=True)
        raise typer.Exit(1)


if __name__ == "__main__":
    typer.run(workbook_create)
