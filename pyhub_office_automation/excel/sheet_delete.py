"""
Excel 워크시트 삭제 명령어 (Typer 버전)
"""

import json
from typing import Optional

import typer

from .utils import ExecutionTimer, create_error_response, create_success_response, get_or_open_workbook


def sheet_delete(
    workbook: Optional[str] = typer.Option(None, "--workbook", help="워크북 파일 경로"),
    use_active: bool = typer.Option(False, "--use-active", help="현재 활성 워크북 사용"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help="열린 워크북 이름으로 접근"),
    name: str = typer.Option(..., "--name", help="삭제할 시트의 이름"),
    force: bool = typer.Option(False, "--force", help="확인 없이 강제 삭제"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택"),
):
    """Excel 워크북에서 시트를 삭제합니다."""
    book = None
    try:
        with ExecutionTimer() as timer:
            book = get_or_open_workbook(file_path=workbook, workbook_name=workbook_name, use_active=use_active, visible=True)

            # 시트 존재 확인
            if name not in [sheet.name for sheet in book.sheets]:
                available_names = [sheet.name for sheet in book.sheets]
                raise ValueError(f"시트 '{name}'을 찾을 수 없습니다. 사용 가능한 시트: {available_names}")

            # 마지막 시트인지 확인
            if len(book.sheets) == 1:
                raise ValueError("마지막 남은 시트는 삭제할 수 없습니다")

            target_sheet = book.sheets[name]
            target_sheet.delete()

            data_content = {
                "deleted_sheet": {"name": name},
                "workbook": {"name": book.name, "remaining_sheets": len(book.sheets)},
            }

            response = create_success_response(
                data=data_content,
                command="sheet-delete",
                message=f"시트 '{name}'을(를) 삭제했습니다",
                execution_time_ms=timer.execution_time_ms,
                book=book,
            )

            if output_format == "json":
                typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
            else:
                typer.echo(f"✅ 시트 '{name}'을(를) 삭제했습니다")

    except Exception as e:
        error_response = create_error_response(e, "sheet-delete")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)


if __name__ == "__main__":
    typer.run(sheet_delete)
