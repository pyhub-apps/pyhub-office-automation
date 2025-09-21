"""
Excel 테이블 쓰기 명령어 (Typer 버전)
"""

import json
from typing import Optional

import pandas as pd
import typer

from .utils import ExecutionTimer, create_error_response, create_success_response, get_or_open_workbook


def table_write(
    workbook: Optional[str] = typer.Option(None, "--workbook", help="워크북 파일 경로"),
    use_active: bool = typer.Option(False, "--use-active", help="현재 활성 워크북 사용"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help="열린 워크북 이름으로 접근"),
    sheet: Optional[str] = typer.Option(None, "--sheet", help="시트 이름"),
    data_file: str = typer.Option(..., "--data-file", help="쓸 데이터 파일 (CSV/JSON)"),
    range_str: str = typer.Option("A1", "--range", help="쓸 시작 위치"),
    header: bool = typer.Option(True, "--header/--no-header", help="헤더 포함 여부"),
    save: bool = typer.Option(True, "--save/--no-save", help="저장 여부"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택"),
):
    """pandas DataFrame을 Excel 테이블로 씁니다."""
    book = None
    try:
        with ExecutionTimer() as timer:
            # 데이터 파일 읽기
            if data_file.endswith(".csv"):
                df = pd.read_csv(data_file)
            elif data_file.endswith(".json"):
                df = pd.read_json(data_file)
            else:
                raise ValueError("지원되지 않는 파일 형식입니다. CSV 또는 JSON 파일을 사용하세요.")

            book = get_or_open_workbook(file_path=workbook, workbook_name=workbook_name, use_active=use_active, visible=False)

            target_sheet = book.sheets.active if not sheet else book.sheets[sheet]
            start_range = target_sheet.range(range_str)

            # DataFrame을 Excel에 쓰기
            if header:
                # 헤더 포함
                values = [df.columns.tolist()] + df.values.tolist()
            else:
                # 헤더 제외
                values = df.values.tolist()

            # 데이터 크기에 맞는 범위 계산
            end_row = start_range.row + len(values) - 1
            end_col = start_range.column + len(values[0]) - 1

            write_range = target_sheet.range((start_range.row, start_range.column), (end_row, end_col))
            write_range.value = values

            if save:
                book.save()

            data_content = {
                "written_data": {"shape": df.shape, "range": write_range.address, "header_included": header},
                "source_file": data_file,
                "saved": save,
            }

            response = create_success_response(
                data=data_content,
                command="table-write",
                message=f"테이블 데이터를 썼습니다 ({df.shape[0]}행 × {df.shape[1]}열)",
                execution_time_ms=timer.execution_time_ms,
                book=book,
            )

            if output_format == "json":
                typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
            else:
                typer.echo(f"✅ 테이블 데이터를 썼습니다 ({df.shape[0]}행 × {df.shape[1]}열)")
                if save:
                    typer.echo("💾 워크북을 저장했습니다")

    except Exception as e:
        error_response = create_error_response(e, "table-write")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)


if __name__ == "__main__":
    typer.run(table_write)
