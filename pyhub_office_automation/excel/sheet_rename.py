"""
Excel 워크시트 이름 변경 명령어 (Typer 버전)
"""

import json
from typing import Optional

import typer

from .utils import ExecutionTimer, create_error_response, create_success_response, get_or_open_workbook


def sheet_rename(
    workbook: Optional[str] = typer.Option(None, "--workbook", help="워크북 파일 경로"),
    use_active: bool = typer.Option(False, "--use-active", help="현재 활성 워크북 사용"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help="열린 워크북 이름으로 접근"),
    old_name: str = typer.Option(..., "--old-name", help="변경할 시트의 현재 이름"),
    new_name: str = typer.Option(..., "--new-name", help="시트의 새 이름"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택"),
):
    """Excel 워크북의 시트 이름을 변경합니다."""
    book = None
    try:
        with ExecutionTimer() as timer:
            book = get_or_open_workbook(file_path=workbook, workbook_name=workbook_name, use_active=use_active, visible=True)

            # 기존 시트 존재 확인
            if old_name not in [sheet.name for sheet in book.sheets]:
                available_names = [sheet.name for sheet in book.sheets]
                raise ValueError(f"시트 '{old_name}'을 찾을 수 없습니다. 사용 가능한 시트: {available_names}")

            # 새 이름 중복 확인
            if new_name in [sheet.name for sheet in book.sheets]:
                raise ValueError(f"시트 이름 '{new_name}'이 이미 존재합니다")

            target_sheet = book.sheets[old_name]
            target_sheet.name = new_name

            data_content = {"renamed_sheet": {"old_name": old_name, "new_name": new_name}, "workbook": {"name": book.name}}

            response = create_success_response(
                data=data_content,
                command="sheet-rename",
                message=f"시트 이름을 '{old_name}'에서 '{new_name}'으로 변경했습니다",
                execution_time_ms=timer.execution_time_ms,
                book=book,
            )

            if output_format == "json":
                typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
            else:
                typer.echo(f"✅ 시트 이름을 '{old_name}'에서 '{new_name}'으로 변경했습니다")

    except Exception as e:
        error_response = create_error_response(e, "sheet-rename")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)


if __name__ == "__main__":
    typer.run(sheet_rename)
