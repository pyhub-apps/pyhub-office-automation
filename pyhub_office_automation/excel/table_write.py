"""
Excel 테이블 쓰기 명령어 (Typer 버전)
pandas DataFrame을 Excel에 쓰고 선택적으로 Excel Table로 변환
"""

import json
import platform
from typing import Optional

import pandas as pd
import typer

from .engines import get_engine
from .utils import ExecutionTimer, create_error_response, create_success_response


def table_write(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="열 Excel 파일의 절대 경로"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help="열린 워크북 이름으로 접근"),
    sheet: Optional[str] = typer.Option(None, "--sheet", help="시트 이름"),
    data_file: str = typer.Option(..., "--data-file", help="쓸 데이터 파일 (CSV/JSON)"),
    range_str: str = typer.Option("A1", "--range", help="쓸 시작 위치"),
    header: bool = typer.Option(True, "--header/--no-header", help="헤더 포함 여부"),
    create_table: bool = typer.Option(
        True, "--create-table/--no-create-table", help="데이터를 Excel Table로 변환 (Windows 전용)"
    ),
    table_name: Optional[str] = typer.Option(None, "--table-name", help="Excel Table 이름 (create-table 사용 시)"),
    table_style: str = typer.Option("TableStyleMedium2", "--table-style", help="Excel Table 스타일"),
    save: bool = typer.Option(True, "--save/--no-save", help="저장 여부"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택"),
    visible: bool = typer.Option(False, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부"),
):
    """
    pandas DataFrame을 Excel에 쓰고 선택적으로 Excel Table로 변환합니다.

    데이터를 Excel에 쓴 후 --create-table 옵션으로 Excel Table을 생성할 수 있습니다.
    Excel Table은 피벗테이블의 동적 범위 확장과 데이터 필터링 기능을 제공합니다.

    \b
    워크북 접근 방법:
      • 옵션 없음: 활성 워크북 자동 사용 (기본값)
      • --file-path: 파일 경로로 워크북 열기
      • --workbook-name: 열린 워크북 이름으로 접근

    \b
    Excel Table 기능:
      • --create-table: 데이터를 Excel Table로 변환 (기본값: True, Windows 전용)
      • --table-name: 테이블 이름 지정 (미지정시 자동 생성)
      • --table-style: 테이블 스타일 선택 (기본값: TableStyleMedium2)

    \b
    사용 예제:
      # 기본 사용 (데이터 쓰기 + Excel Table 생성)
      oa excel table-write --data-file "data.csv"

      # Excel Table 없이 데이터만 쓰기
      oa excel table-write --data-file "data.csv" --no-create-table

      # 커스텀 테이블 설정
      oa excel table-write --data-file "data.csv" --table-name "SalesData" --table-style "TableStyleMedium5"

      # 특정 위치에 쓰기
      oa excel table-write --data-file "data.csv" --range "C3" --table-name "CustomTable"
    """
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

            # Engine 획득
            engine = get_engine()

            # 워크북 연결
            if file_path:
                book = engine.open_workbook(file_path, visible=visible)
            elif workbook_name:
                book = engine.get_workbook_by_name(workbook_name)
            else:
                book = engine.get_active_workbook()

            # 대상 시트 결정 (COM API 직접 사용)
            target_sheet = book.ActiveSheet if not sheet else book.Sheets(sheet)
            start_range = target_sheet.Range(range_str)

            # DataFrame을 Excel에 쓰기
            if header:
                # 헤더 포함
                values = [df.columns.tolist()] + df.values.tolist()
            else:
                # 헤더 제외
                values = df.values.tolist()

            # 데이터 크기에 맞는 범위 계산 (COM API 사용)
            end_row = start_range.Row + len(values) - 1
            end_col = start_range.Column + len(values[0]) - 1

            write_range = target_sheet.Range(start_range, target_sheet.Cells(end_row, end_col))
            write_range.Value = values

            # Excel Table 생성 (옵션이 활성화된 경우)
            table_info = None
            if create_table:
                if platform.system() != "Windows":
                    # macOS에서는 경고만 표시하고 계속 진행
                    table_info = {"warning": "Excel Table 생성은 Windows에서만 지원됩니다."}
                else:
                    try:
                        # 테이블 이름 자동 생성 (COM API 사용)
                        if not table_name:
                            existing_tables = []
                            for lo in target_sheet.ListObjects:
                                existing_tables.append(lo.Name)
                            counter = 1
                            while True:
                                candidate_name = f"Table{counter}"
                                if candidate_name not in existing_tables:
                                    table_name = candidate_name
                                    break
                                counter += 1

                        # 테이블 이름 중복 확인
                        existing_table_names = []
                        for lo in target_sheet.ListObjects:
                            existing_table_names.append(lo.Name)
                        if table_name in existing_table_names:
                            # 중복 시 숫자 suffix 추가
                            base_name = table_name
                            counter = 2
                            while table_name in existing_table_names:
                                table_name = f"{base_name}{counter}"
                                counter += 1

                        # Excel Table 생성 (Windows COM API 사용)
                        list_object = target_sheet.ListObjects.Add(
                            SourceType=1,  # xlSrcRange
                            Source=write_range,
                            XlListObjectHasHeaders=1 if header else 2,  # xlYes=1, xlNo=2
                        )

                        # 테이블 이름 설정
                        list_object.Name = table_name

                        # 테이블 스타일 적용
                        try:
                            list_object.TableStyle = table_style
                        except:
                            # 스타일 적용 실패 시 기본 스타일 사용
                            list_object.TableStyle = "TableStyleMedium2"
                            table_style = "TableStyleMedium2"

                        table_info = {
                            "name": table_name,
                            "range": write_range.Address,
                            "style": table_style,
                            "has_headers": header,
                            "created": True,
                        }

                    except Exception as e:
                        table_info = {"error": f"Excel Table 생성 실패: {str(e)}"}

            # 저장 처리 (COM API 사용)
            saved = False
            if save:
                try:
                    book.Save()
                    saved = True
                except Exception as e:
                    # 저장 실패해도 데이터는 쓰여진 상태
                    pass

            data_content = {
                "written_data": {"shape": df.shape, "range": write_range.Address, "header_included": header},
                "table": table_info,
                "source_file": data_file,
                "saved": saved,
            }

            # 성공 메시지 생성
            table_status = ""
            if create_table and table_info:
                if "created" in table_info and table_info["created"]:
                    table_status = f", Excel Table '{table_info['name']}' 생성됨"
                elif "warning" in table_info:
                    table_status = f", {table_info['warning']}"
                elif "error" in table_info:
                    table_status = f", {table_info['error']}"

            save_status = "저장됨" if saved else ("저장하지 않음" if not save else "저장 실패")
            message = f"테이블 데이터를 썼습니다 ({df.shape[0]}행 × {df.shape[1]}열{table_status}, {save_status})"

            response = create_success_response(
                data=data_content,
                command="table-write",
                message=message,
                execution_time_ms=timer.execution_time_ms,
                book=book,
                range_obj=write_range,
            )

            if output_format == "json":
                typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
            else:
                typer.echo(f"✅ {message}")

                if table_info:
                    if "created" in table_info and table_info["created"]:
                        typer.echo(f"🏷️ Excel Table: {table_info['name']} ({table_info['style']})")
                    elif "warning" in table_info:
                        typer.echo(f"⚠️ {table_info['warning']}")
                    elif "error" in table_info:
                        typer.echo(f"❌ {table_info['error']}")

                if saved:
                    typer.echo("💾 워크북을 저장했습니다")
                elif not save:
                    typer.echo("⚠️ 저장하지 않음 (--no-save 옵션)")
                else:
                    typer.echo("❌ 저장 실패")

    except Exception as e:
        error_response = create_error_response(e, "table-write")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)

    finally:
        # 워크북 정리 - 파일 경로로 열었고 visible=False인 경우에만 앱 종료
        if book is not None and not visible and file_path:
            try:
                book.Application.Quit()
            except:
                pass


if __name__ == "__main__":
    typer.run(table_write)
