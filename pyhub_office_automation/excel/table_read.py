"""
Excel 테이블 읽기 명령어 (Typer 버전)
"""

import json
import platform
from typing import Optional

import pandas as pd
import typer

from .utils import ExecutionTimer, create_error_response, create_success_response, get_or_open_workbook


def table_read(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="열 Excel 파일의 절대 경로"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help="열린 워크북 이름으로 접근"),
    sheet: Optional[str] = typer.Option(None, "--sheet", help="시트 이름"),
    table_name: Optional[str] = typer.Option(None, "--table-name", help="테이블 이름"),
    range_str: Optional[str] = typer.Option(None, "--range", help="읽을 테이블 범위"),
    header: bool = typer.Option(True, "--header/--no-header", help="첫 행을 헤더로 사용"),
    offset: Optional[int] = typer.Option(None, "--offset", help="시작 행 번호 (0부터)"),
    limit: Optional[int] = typer.Option(None, "--limit", help="읽을 행 수"),
    sample_mode: bool = typer.Option(False, "--sample-mode", help="지능형 샘플링 모드 (첫/중간/마지막)"),
    columns: Optional[str] = typer.Option(None, "--columns", help="읽을 컬럼명 (쉼표로 구분)"),
    output_file: Optional[str] = typer.Option(None, "--output-file", help="결과를 저장할 CSV 파일"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택"),
):
    """Excel 테이블 데이터를 pandas DataFrame으로 읽습니다."""
    book = None
    try:
        with ExecutionTimer() as timer:
            book = get_or_open_workbook(file_path=file_path, workbook_name=workbook_name, visible=False)

            target_sheet = book.sheets.active if not sheet else book.sheets[sheet]

            if range_str:
                # 지정된 범위에서 읽기
                range_obj = target_sheet.range(range_str)
                values = range_obj.value
            elif table_name:
                # 테이블 이름으로 읽기
                target_table = None
                target_table_sheet = None

                if sheet:
                    # 특정 시트에서 테이블 찾기
                    for table in target_sheet.tables:
                        if table.name == table_name:
                            target_table = table
                            target_table_sheet = target_sheet
                            break
                else:
                    # 모든 시트에서 테이블 찾기
                    for sheet_obj in book.sheets:
                        for table in sheet_obj.tables:
                            if table.name == table_name:
                                target_table = table
                                target_table_sheet = sheet_obj
                                break
                        if target_table:
                            break

                if not target_table:
                    raise ValueError(f"테이블 '{table_name}'을(를) 찾을 수 없습니다")

                # 테이블 데이터 읽기
                table_range = target_table.range
                all_values = table_range.value

                # 헤더와 데이터 분리
                if isinstance(all_values, list) and len(all_values) > 0:
                    if header and len(all_values) > 1:
                        headers = all_values[0]
                        data = all_values[1:]
                    else:
                        headers = None
                        data = all_values

                    # 컬럼 선택
                    if columns and headers:
                        selected_cols = [col.strip() for col in columns.split(',')]
                        col_indices = []
                        selected_headers = []
                        for col in selected_cols:
                            if col in headers:
                                col_indices.append(headers.index(col))
                                selected_headers.append(col)

                        if col_indices:
                            headers = selected_headers
                            data = [[row[i] if i < len(row) else None for i in col_indices] for row in data]

                    # 오프셋과 제한 적용
                    total_rows = len(data)
                    start_idx = offset if offset else 0

                    if start_idx >= total_rows:
                        data = []
                    else:
                        if limit:
                            if sample_mode and total_rows > limit:
                                # 지능형 샘플링: 첫 20%, 중간 60%, 마지막 20%
                                first_count = max(1, int(limit * 0.2))
                                last_count = max(1, int(limit * 0.2))
                                middle_count = limit - first_count - last_count

                                sampled_data = []
                                # 첫 부분
                                sampled_data.extend(data[:first_count])

                                # 중간 부분
                                if middle_count > 0 and total_rows > first_count + last_count:
                                    middle_start = first_count
                                    middle_end = total_rows - last_count
                                    middle_indices = range(middle_start, middle_end,
                                                         max(1, (middle_end - middle_start) // middle_count))
                                    sampled_data.extend([data[i] for i in middle_indices[:middle_count]])

                                # 마지막 부분
                                if last_count > 0 and total_rows > last_count:
                                    sampled_data.extend(data[-last_count:])

                                data = sampled_data
                            else:
                                # 일반 제한
                                end_idx = start_idx + limit
                                data = data[start_idx:end_idx]
                        else:
                            # 오프셋만 적용
                            data = data[start_idx:]

                    # 최종 values 구성
                    if headers and header:
                        values = [headers] + data
                    else:
                        values = data
                else:
                    values = []

                # 현재 시트를 테이블이 있는 시트로 변경
                if target_table_sheet != target_sheet:
                    target_sheet = target_table_sheet
            else:
                # table_name도 range_str도 없는 경우: 더 유용한 안내 제공
                # 워크북의 모든 테이블 정보 수집
                all_tables = []
                for sheet_obj in book.sheets:
                    for table in sheet_obj.tables:
                        all_tables.append(f"'{table.name}' (시트: {sheet_obj.name})")

                if all_tables:
                    tables_str = ", ".join(all_tables)
                    # 현재 시트에 테이블이 있는지 확인
                    sheet_tables = [table.name for table in target_sheet.tables]
                    if sheet_tables:
                        table_list_str = ", ".join(f"'{name}'" for name in sheet_tables)
                        raise ValueError(
                            f"테이블 이름을 지정해주세요. "
                            f"현재 시트({target_sheet.name}) 테이블: {table_list_str} | "
                            f"모든 테이블: {tables_str}"
                        )
                    else:
                        raise ValueError(
                            f"현재 시트({target_sheet.name})에 테이블이 없습니다. "
                            f"사용 가능한 테이블: {tables_str} | "
                            f"또는 --range 옵션을 사용하세요."
                        )

                # 테이블이 없으면 used_range로 읽기 시도 (후순위)
                used_range = target_sheet.used_range
                if not used_range:
                    raise ValueError(f"시트({target_sheet.name})에 데이터가 없습니다. --table-name 또는 --range 옵션을 사용하세요.")

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

            # JSON 직렬화 가능한 preview 데이터 생성
            preview_data = []

            if not df.empty:
                # limit 지정 여부에 관계없이 전체 데이터를 preview로 제공
                # (table-read는 데이터 조회 명령어이므로 전체 반환이 기본)
                for record in df.to_dict("records"):
                    clean_record = {}
                    for key, value in record.items():
                        if pd.isna(value) or value is None:
                            clean_record[key] = None
                        elif isinstance(value, (str, int, float, bool)):
                            clean_record[key] = value
                        else:
                            clean_record[key] = str(value)
                    preview_data.append(clean_record)

            data_content = {
                "dataframe_info": {
                    "shape": df.shape,
                    "columns": df.columns.tolist() if not df.empty else [],
                    "dtypes": {col: str(dtype) for col, dtype in df.dtypes.items()} if not df.empty else {},
                },
                "data": preview_data,  # "preview" → "data"로 명칭 변경 (전체 데이터이므로)
                "output_file": output_file,
            }

            # 테이블 읽기 추가 정보
            if table_name:
                data_content.update({
                    "table_name": table_name,
                    "sheet": target_sheet.name,
                    "offset": offset if offset else 0,
                    "limit": limit,
                    "sample_mode": sample_mode,
                    "selected_columns": columns.split(',') if columns else None,
                })

            response = create_success_response(
                data=data_content,
                command="table-read",
                message=f"테이블 데이터를 읽었습니다 ({df.shape[0]}행 × {df.shape[1]}열)",
                execution_time_ms=timer.execution_time_ms,
                book=book,
            )

            if output_format == "json":
                typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
            elif output_format == "csv":
                # CSV 형식으로 데이터 출력
                if not df.empty:
                    typer.echo(df.to_csv(index=False))
                else:
                    typer.echo("# 데이터가 없습니다")
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
