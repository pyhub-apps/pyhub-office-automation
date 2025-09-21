"""
Excel 워크시트 추가 명령어 (Typer 버전)
"""

import json
from typing import Optional

import typer

from pyhub_office_automation.version import get_version

from .utils import ExecutionTimer, create_error_response, create_success_response, get_or_open_workbook


def sheet_add(
    workbook: Optional[str] = typer.Option(None, "--workbook", help="워크북 파일 경로"),
    use_active: bool = typer.Option(False, "--use-active", help="현재 활성 워크북 사용"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help="열린 워크북 이름으로 접근"),
    name: str = typer.Option(..., "--name", help="추가할 시트의 이름"),
    before: Optional[str] = typer.Option(None, "--before", help="이 시트 앞에 추가"),
    after: Optional[str] = typer.Option(None, "--after", help="이 시트 뒤에 추가"),
    visible: bool = typer.Option(True, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택"),
):
    """Excel 워크북에 새 워크시트를 추가합니다."""
    book = None
    try:
        with ExecutionTimer() as timer:
            book = get_or_open_workbook(
                file_path=workbook, workbook_name=workbook_name, use_active=use_active, visible=visible
            )

            # 기존 시트명 중복 확인
            existing_names = [sheet.name for sheet in book.sheets]
            if name in existing_names:
                raise ValueError(f"시트 '{name}'이 이미 존재합니다")

            # 시트 추가
            if before:
                new_sheet = book.sheets.add(name, before=before)
            elif after:
                new_sheet = book.sheets.add(name, after=after)
            else:
                new_sheet = book.sheets.add(name)

            data_content = {
                "added_sheet": {"name": new_sheet.name, "index": new_sheet.index},
                "workbook": {"name": book.name, "total_sheets": len(book.sheets)},
            }

            response = create_success_response(
                data=data_content,
                command="sheet-add",
                message=f"시트 '{name}'을(를) 추가했습니다",
                execution_time_ms=timer.execution_time_ms,
                book=book,
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
