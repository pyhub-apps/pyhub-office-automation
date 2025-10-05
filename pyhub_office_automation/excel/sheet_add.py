"""
Excel 워크시트 추가 명령어 (Engine 기반)
"""

import json
from typing import Optional

import typer

from pyhub_office_automation.version import get_version

from .engines import get_engine
from .utils import ExecutionTimer, create_error_response, create_success_response


def sheet_add(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="열 Excel 파일의 절대 경로"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help="열린 워크북 이름으로 접근"),
    name: str = typer.Option(..., "--name", help="추가할 시트의 이름"),
    before: Optional[str] = typer.Option(None, "--before", help="이 시트 앞에 추가"),
    after: Optional[str] = typer.Option(None, "--after", help="이 시트 뒤에 추가"),
    visible: bool = typer.Option(True, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택"),
):
    """Excel 워크북에 새 워크시트를 추가합니다."""
    try:
        # before와 after는 동시에 사용 불가
        if before and after:
            raise ValueError("--before와 --after는 동시에 사용할 수 없습니다")

        with ExecutionTimer() as timer:
            # Engine 획득
            engine = get_engine()

            # 워크북 가져오기
            if file_path:
                book = engine.open_workbook(file_path, visible=visible)
            elif workbook_name:
                book = engine.get_workbook_by_name(workbook_name)
            else:
                book = engine.get_active_workbook()

            # 기존 워크북 정보 가져오기
            wb_info_before = engine.get_workbook_info(book)
            existing_sheets = wb_info_before["sheets"]

            # 중복 확인
            if name in existing_sheets:
                raise ValueError(f"시트 '{name}'이 이미 존재합니다")

            # before 옵션 사용 시 (Engine의 add_sheet는 before만 지원)
            position_arg = None
            if before:
                if before not in existing_sheets:
                    raise ValueError(f"시트 '{before}'를 찾을 수 없습니다")
                position_arg = before
            elif after:
                # after는 xlwings 전용이므로 변환 필요
                # after "Sheet1" → before "Sheet2" (다음 시트)
                if after not in existing_sheets:
                    raise ValueError(f"시트 '{after}'를 찾을 수 없습니다")
                after_index = existing_sheets.index(after)
                if after_index + 1 < len(existing_sheets):
                    position_arg = existing_sheets[after_index + 1]
                # 마지막 시트 뒤에 추가하는 경우 before=None (맨 뒤)

            # Engine을 통해 시트 추가
            new_sheet_name = engine.add_sheet(book, name, before=position_arg)

            # 추가 후 워크북 정보 가져오기
            wb_info_after = engine.get_workbook_info(book)
            new_sheets = wb_info_after["sheets"]
            new_sheet_index = new_sheets.index(new_sheet_name) if new_sheet_name in new_sheets else -1

            data_content = {
                "added_sheet": {"name": new_sheet_name, "index": new_sheet_index},
                "workbook": {"name": wb_info_after["name"], "total_sheets": wb_info_after["sheet_count"]},
            }

            response = create_success_response(
                data=data_content,
                command="sheet-add",
                message=f"시트 '{name}'을(를) 추가했습니다",
                execution_time_ms=timer.execution_time_ms,
            )

            if output_format == "json":
                typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
            else:
                typer.echo(f"✅ 시트 '{name}'을(를) 추가했습니다")

    except Exception as e:
        error_response = create_error_response(e, "sheet-add")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)


if __name__ == "__main__":
    typer.run(sheet_add)
