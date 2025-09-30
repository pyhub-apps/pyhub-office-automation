"""
Excel 워크시트 이름 변경 명령어 (Typer 버전)
"""

import json
from typing import Optional

import typer

from .utils import ExecutionTimer, create_error_response, create_success_response, get_or_open_workbook


def sheet_rename(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="열 Excel 파일의 절대 경로"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help="열린 워크북 이름으로 접근"),
    old_sheet: Optional[str] = typer.Option(None, "--old-sheet", help="변경할 시트의 현재 이름"),
    new_sheet: Optional[str] = typer.Option(None, "--new-sheet", help="시트의 새 이름"),
    old_name: Optional[str] = typer.Option(None, "--old-name", help="[별칭] 변경할 시트의 현재 이름 (--old-sheet 사용 권장)"),
    new_name: Optional[str] = typer.Option(None, "--new-name", help="[별칭] 시트의 새 이름 (--new-sheet 사용 권장)"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택"),
):
    """Excel 워크북의 시트 이름을 변경합니다."""
    book = None
    try:
        # 옵션 우선순위 처리 (새 옵션 우선)
        current_name = old_sheet or old_name
        target_name = new_sheet or new_name

        if not current_name or not target_name:
            raise ValueError("--old-sheet(또는 --old-name)와 --new-sheet(또는 --new-name) 옵션을 모두 지정해야 합니다")

        with ExecutionTimer() as timer:
            book = get_or_open_workbook(file_path=file_path, workbook_name=workbook_name, visible=True)

            # 기존 시트 존재 확인
            if current_name not in [sheet.name for sheet in book.sheets]:
                available_names = [sheet.name for sheet in book.sheets]
                raise ValueError(f"시트 '{current_name}'을 찾을 수 없습니다. 사용 가능한 시트: {available_names}")

            # 새 이름 중복 확인
            if target_name in [sheet.name for sheet in book.sheets]:
                raise ValueError(f"시트 이름 '{target_name}'이 이미 존재합니다")

            target_sheet = book.sheets[current_name]
            target_sheet.name = target_name

            data_content = {
                "renamed_sheet": {"old_name": current_name, "new_name": target_name},
                "workbook": {"name": book.name},
            }

            response = create_success_response(
                data=data_content,
                command="sheet-rename",
                message=f"시트 이름을 '{current_name}'에서 '{target_name}'으로 변경했습니다",
                execution_time_ms=timer.execution_time_ms,
                book=book,
            )

            if output_format == "json":
                typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
            else:
                typer.echo(f"✅ 시트 이름을 '{current_name}'에서 '{target_name}'으로 변경했습니다")

    except Exception as e:
        error_response = create_error_response(e, "sheet-rename")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)


if __name__ == "__main__":
    typer.run(sheet_rename)
