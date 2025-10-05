"""
Excel 워크시트 삭제 명령어 (Engine 기반)
"""

import json
from typing import Optional

import typer

from .engines import get_engine
from .utils import ExecutionTimer, create_error_response, create_success_response


def sheet_delete(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="열 Excel 파일의 절대 경로"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help="열린 워크북 이름으로 접근"),
    sheet: Optional[str] = typer.Option(None, "--sheet", help="삭제할 시트의 이름"),
    name: Optional[str] = typer.Option(None, "--name", help="[별칭] 삭제할 시트의 이름 (--sheet 사용 권장)"),
    force: bool = typer.Option(False, "--force", help="확인 없이 강제 삭제"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택"),
):
    """Excel 워크북에서 시트를 삭제합니다."""
    try:
        # 옵션 우선순위 처리 (새 옵션 우선)
        sheet_name = sheet or name
        if not sheet_name:
            raise ValueError("--sheet(또는 --name) 옵션으로 삭제할 시트를 지정해야 합니다")

        with ExecutionTimer() as timer:
            # Engine 획득
            engine = get_engine()

            # 워크북 가져오기
            if file_path:
                book = engine.open_workbook(file_path, visible=True)
            elif workbook_name:
                book = engine.get_workbook_by_name(workbook_name)
            else:
                book = engine.get_active_workbook()

            # 워크북 정보 가져오기
            wb_info = engine.get_workbook_info(book)
            existing_sheets = wb_info["sheets"]

            # 시트 존재 확인
            if sheet_name not in existing_sheets:
                raise ValueError(f"시트 '{sheet_name}'을 찾을 수 없습니다. 사용 가능한 시트: {existing_sheets}")

            # 마지막 시트인지 확인
            if len(existing_sheets) == 1:
                raise ValueError("마지막 남은 시트는 삭제할 수 없습니다")

            # Engine을 통해 시트 삭제
            engine.delete_sheet(book, sheet_name)

            # 삭제 후 워크북 정보 가져오기
            wb_info_after = engine.get_workbook_info(book)

            data_content = {
                "deleted_sheet": {"name": sheet_name},
                "workbook": {"name": wb_info_after["name"], "remaining_sheets": wb_info_after["sheet_count"]},
            }

            response = create_success_response(
                data=data_content,
                command="sheet-delete",
                message=f"시트 '{sheet_name}'을(를) 삭제했습니다",
                execution_time_ms=timer.execution_time_ms,
            )

            if output_format == "json":
                typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
            else:
                typer.echo(f"✅ 시트 '{sheet_name}'을(를) 삭제했습니다")

    except Exception as e:
        error_response = create_error_response(e, "sheet-delete")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)


if __name__ == "__main__":
    typer.run(sheet_delete)
