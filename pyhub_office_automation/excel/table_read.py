"""
Excel 테이블 읽기 명령어 (Typer 버전)
"""

import json
from typing import Optional

import pandas as pd
import typer

from .utils import ExecutionTimer, create_error_response, create_success_response, get_or_open_workbook


def table_read(
    workbook: Optional[str] = typer.Option(None, "--workbook", help="워크북 파일 경로"),
    use_active: bool = typer.Option(False, "--use-active", help="현재 활성 워크북 사용"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help="열린 워크북 이름으로 접근"),
    sheet: Optional[str] = typer.Option(None, "--sheet", help="시트 이름"),
    table_name: Optional[str] = typer.Option(None, "--table-name", help="테이블 이름"),
    range_str: Optional[str] = typer.Option(None, "--range", help="읽을 테이블 범위"),
    header: bool = typer.Option(True, "--header/--no-header", help="첫 행을 헤더로 사용"),
    output_file: Optional[str] = typer.Option(None, "--output-file", help="결과를 저장할 CSV 파일"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택"),
):
    """Excel 테이블 데이터를 pandas DataFrame으로 읽습니다."""
    book = None
    try:
        with ExecutionTimer() as timer:
            book = get_or_open_workbook(file_path=workbook, workbook_name=workbook_name, use_active=use_active, visible=False)

            target_sheet = book.sheets.active if not sheet else book.sheets[sheet]

            if range_str:
                # 지정된 범위에서 읽기
                range_obj = target_sheet.range(range_str)
                values = range_obj.value
            elif table_name:
                # 테이블 이름으로 읽기 (구현 필요)
                raise NotImplementedError("테이블 이름으로 읽기는 아직 구현되지 않았습니다")
            else:
                # 사용된 범위 전체 읽기
                used_range = target_sheet.used_range
                if not used_range:
                    raise ValueError("시트에 데이터가 없습니다")
                values = used_range.value

            # pandas DataFrame 생성
            if isinstance(values, list) and len(values) > 0:
                if header and len(values) > 1:
                    df = pd.DataFrame(values[1:], columns=values[0])
                else:
                    df = pd.DataFrame(values)
            else:
                df = pd.DataFrame()

            # 출력 파일 저장
            if output_file:
                df.to_csv(output_file, index=False)

            data_content = {
                "dataframe_info": {
                    "shape": df.shape,
                    "columns": df.columns.tolist() if not df.empty else [],
                    "dtypes": df.dtypes.to_dict() if not df.empty else {},
                },
                "preview": df.head().to_dict("records") if not df.empty else [],
                "output_file": output_file,
            }

            response = create_success_response(
                data=data_content,
                command="table-read",
                message=f"테이블 데이터를 읽었습니다 ({df.shape[0]}행 × {df.shape[1]}열)",
                execution_time_ms=timer.execution_time_ms,
                book=book,
            )

            if output_format == "json":
                typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
            else:
                typer.echo(f"✅ 테이블 데이터를 읽었습니다 ({df.shape[0]}행 × {df.shape[1]}열)")
                if output_file:
                    typer.echo(f"💾 결과를 '{output_file}'에 저장했습니다")

    except Exception as e:
        error_response = create_error_response(e, "table-read")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)


if __name__ == "__main__":
    typer.run(table_read)
