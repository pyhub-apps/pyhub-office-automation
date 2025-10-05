"""
Excel 테이블 읽기 명령어 (Typer 버전)
"""

import json
import platform
from typing import Optional

import pandas as pd
import typer

from .engines import get_engine
from .utils import ExecutionTimer, create_error_response, create_success_response


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
            # Engine 획득
            engine = get_engine()

            # 워크북 연결
            if file_path:
                book = engine.open_workbook(file_path, visible=False)
            elif workbook_name:
                book = engine.get_workbook_by_name(workbook_name)
            else:
                book = engine.get_active_workbook()

            # 대상 시트 결정 (COM API 직접 사용)
            target_sheet = book.ActiveSheet if not sheet else book.Sheets(sheet)

            if range_str:
                # 지정된 범위에서 읽기 (COM API 직접 사용)
                range_obj = target_sheet.Range(range_str)
                values = range_obj.Value
                # COM API는 단일 셀을 스칼라로 반환하므로 리스트로 변환
                if not isinstance(values, (list, tuple)):
                    values = [[values]]
                elif values and not isinstance(values[0], (list, tuple)):
                    values = [values]
            elif table_name:
                # 테이블 이름으로 읽기 (Engine 메서드 사용)
                col_list = [col.strip() for col in columns.split(",")] if columns else None

                # Engine의 read_table()은 offset과 limit을 지원하지만 sample_mode는 지원하지 않음
                # sample_mode인 경우 전체 데이터를 가져온 후 직접 샘플링
                if sample_mode and limit:
                    table_result = engine.read_table(book, table_name, columns=col_list, offset=offset or 0)
                else:
                    table_result = engine.read_table(book, table_name, columns=col_list, limit=limit, offset=offset or 0)

                headers = table_result["headers"]
                data = table_result["data"]

                # 샘플링 모드 처리
                if sample_mode and limit and len(data) > limit:
                    # 지능형 샘플링: 첫 20%, 중간 60%, 마지막 20%
                    first_count = max(1, int(limit * 0.2))
                    last_count = max(1, int(limit * 0.2))
                    middle_count = limit - first_count - last_count

                    sampled_data = []
                    # 첫 부분
                    sampled_data.extend(data[:first_count])

                    # 중간 부분
                    total_rows = len(data)
                    if middle_count > 0 and total_rows > first_count + last_count:
                        middle_start = first_count
                        middle_end = total_rows - last_count
                        middle_indices = range(middle_start, middle_end, max(1, (middle_end - middle_start) // middle_count))
                        sampled_data.extend([data[i] for i in middle_indices[:middle_count]])

                    # 마지막 부분
                    if last_count > 0 and total_rows > last_count:
                        sampled_data.extend(data[-last_count:])

                    data = sampled_data

                # 최종 values 구성
                if headers and header:
                    values = [headers] + data
                else:
                    values = data

                # 테이블이 있는 시트 이름 가져오기 (COM API 사용)
                for ws in book.Sheets:
                    try:
                        ws.ListObjects(table_name)
                        target_sheet = ws
                        break
                    except:
                        continue
            else:
                # table_name도 range_str도 없는 경우: Engine을 사용해 모든 테이블 정보 수집
                all_table_infos = engine.list_tables(book)
                all_tables = [f"'{t.name}' (시트: {t.sheet_name})" for t in all_table_infos]

                if all_tables:
                    tables_str = ", ".join(all_tables)
                    # 현재 시트에 테이블이 있는지 확인
                    sheet_tables = [t.name for t in all_table_infos if t.sheet_name == target_sheet.Name]
                    if sheet_tables:
                        table_list_str = ", ".join(f"'{name}'" for name in sheet_tables)
                        raise ValueError(
                            f"테이블 이름을 지정해주세요. "
                            f"현재 시트({target_sheet.Name}) 테이블: {table_list_str} | "
                            f"모든 테이블: {tables_str}"
                        )
                    else:
                        raise ValueError(
                            f"현재 시트({target_sheet.Name})에 테이블이 없습니다. "
                            f"사용 가능한 테이블: {tables_str} | "
                            f"또는 --range 옵션을 사용하세요."
                        )

                # 테이블이 없으면 used_range로 읽기 시도 (후순위) - COM API 사용
                used_range = target_sheet.UsedRange
                if not used_range:
                    raise ValueError(
                        f"시트({target_sheet.Name})에 데이터가 없습니다. --table-name 또는 --range 옵션을 사용하세요."
                    )

                values = used_range.Value
                # COM API 값 정규화
                if not isinstance(values, (list, tuple)):
                    values = [[values]]
                elif values and not isinstance(values[0], (list, tuple)):
                    values = [values]

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
                data_content.update(
                    {
                        "table_name": table_name,
                        "sheet": target_sheet.Name,
                        "offset": offset if offset else 0,
                        "limit": limit,
                        "sample_mode": sample_mode,
                        "selected_columns": columns.split(",") if columns else None,
                    }
                )

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

    finally:
        # 워크북 정리 - 파일 경로로 열었고 visible=False인 경우에만 앱 종료
        if book and file_path:
            try:
                book.Application.Quit()
            except:
                pass


if __name__ == "__main__":
    typer.run(table_read)
