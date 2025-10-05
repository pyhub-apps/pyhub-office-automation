"""
Excel 테이블 정렬 명령어 (Typer 버전)
Excel Table(ListObject)에 단일 또는 다중 컬럼 정렬 적용
"""

import json
import platform
from pathlib import Path
from typing import Optional

import typer
import xlwings as xw

from pyhub_office_automation.version import get_version

from .engines import get_engine
from .utils import (
    ExecutionTimer,
    create_error_response,
    create_success_response,
    get_or_open_workbook,
    get_sheet,
    normalize_path,
)


def table_sort(
    table_name: str = typer.Option(..., "--table-name", help="정렬할 테이블 이름"),
    file_path: Optional[str] = typer.Option(None, "--file-path", help="Excel 파일의 절대 경로"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help="열린 워크북 이름으로 접근"),
    sheet: Optional[str] = typer.Option(None, "--sheet", help="시트 이름 (미지정시 테이블 이름으로 검색)"),
    column: Optional[str] = typer.Option(None, "--column", help="정렬 기준 컬럼 (단일 컬럼, --columns와 동시 사용 불가)"),
    columns: Optional[str] = typer.Option(None, "--columns", help="정렬 기준 컬럼들 (콤마로 구분, 최대 3개)"),
    order: str = typer.Option("asc", "--order", help="정렬 순서 (asc/desc, --column 사용시만 적용)"),
    orders: Optional[str] = typer.Option(None, "--orders", help="각 컬럼별 정렬 순서 (콤마로 구분, --columns와 함께 사용)"),
    save: bool = typer.Option(True, "--save/--no-save", help="정렬 후 파일 저장 여부"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택"),
    visible: bool = typer.Option(False, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부"),
):
    """
    Excel Table에 단일 또는 다중 컬럼 정렬을 적용합니다.

    정렬 기능은 Excel의 AutoFilter.Sort 메서드를 사용하여 구현되며,
    Windows 전용 기능입니다. macOS에서는 에러가 발생합니다.

    \b
    워크북 접근 방법:
      • 옵션 없음: 활성 워크북 자동 사용 (기본값)
      • --file-path: 파일 경로로 워크북 열기
      • --workbook-name: 열린 워크북 이름으로 접근

    \b
    정렬 옵션:
      • 단일 컬럼: --column "ColumnName" --order "asc|desc"
      • 다중 컬럼: --columns "Col1,Col2,Col3" --orders "asc,desc,asc"
      • 최대 3개 컬럼까지 지원 (Excel API 제한)

    \b
    정렬 순서:
      • asc: 오름차순 (기본값)
      • desc: 내림차순

    \b
    사용 예제:
      # 단일 컬럼 정렬 (Amount 컬럼을 내림차순으로)
      oa excel table-sort --table-name "SalesData" --column "Amount" --order "desc"

      # 다중 컬럼 정렬 (Date 오름차순, Amount 내림차순)
      oa excel table-sort --table-name "SalesData" --columns "Date,Amount" --orders "asc,desc"

      # 특정 시트의 테이블 정렬
      oa excel table-sort --table-name "ProductTable" --sheet "Products" --column "Price" --order "asc"

      # 저장하지 않고 정렬만 적용
      oa excel table-sort --table-name "TempData" --column "Name" --no-save
    """
    book = None
    try:
        with ExecutionTimer() as timer:
            # 플랫폼 확인
            if platform.system() != "Windows":
                raise ValueError("Excel Table 정렬은 Windows에서만 지원됩니다.")

            # 옵션 검증
            if column and columns:
                raise typer.BadParameter("--column과 --columns 옵션은 동시에 사용할 수 없습니다.")

            if not column and not columns:
                raise typer.BadParameter("--column 또는 --columns 옵션 중 하나는 필수입니다.")

            # 정렬 설정 파싱
            sort_configs = []

            if column:
                # 단일 컬럼 정렬
                if order.lower() not in ["asc", "desc"]:
                    raise typer.BadParameter("--order는 'asc' 또는 'desc'만 가능합니다.")
                sort_configs.append({"column": column.strip(), "order": order.lower()})
            else:
                # 다중 컬럼 정렬
                column_list = [col.strip() for col in columns.split(",")]
                if len(column_list) > 3:
                    raise typer.BadParameter("최대 3개의 컬럼까지만 정렬할 수 있습니다.")

                if orders:
                    order_list = [ord.strip().lower() for ord in orders.split(",")]
                    if len(order_list) != len(column_list):
                        raise typer.BadParameter("컬럼 수와 정렬 순서 수가 일치하지 않습니다.")

                    for order_item in order_list:
                        if order_item not in ["asc", "desc"]:
                            raise typer.BadParameter("정렬 순서는 'asc' 또는 'desc'만 가능합니다.")
                else:
                    order_list = ["asc"] * len(column_list)

                for col, ord in zip(column_list, order_list):
                    sort_configs.append({"column": col, "order": ord})

            # 워크북 연결
            book = get_or_open_workbook(file_path=file_path, workbook_name=workbook_name, visible=visible)

            # 테이블 찾기
            target_table = None
            target_sheet = None

            if sheet:
                # 특정 시트에서 테이블 찾기
                target_sheet = get_sheet(book, sheet)
                for table in target_sheet.tables:
                    if table.name == table_name:
                        target_table = table
                        break
            else:
                # 모든 시트에서 테이블 찾기
                for sheet_obj in book.sheets:
                    for table in sheet_obj.tables:
                        if table.name == table_name:
                            target_table = table
                            target_sheet = sheet_obj
                            break
                    if target_table:
                        break

            if not target_table:
                sheet_msg = f"시트 '{sheet}'" if sheet else "워크북"
                raise ValueError(f"{sheet_msg}에서 테이블 '{table_name}'을 찾을 수 없습니다.")

            # Engine 가져오기
            engine = get_engine()

            # 정렬 필드를 Engine에 맞는 형식으로 준비
            engine_sort_fields = []
            for config in sort_configs:
                engine_sort_fields.append({"column": config["column"], "order": config["order"]})

            # 정렬 적용 (Engine Layer 사용)
            try:
                result = engine.sort_table(
                    workbook=book.api, sheet_name=target_sheet.name, table_name=table_name, sort_fields=engine_sort_fields
                )

                # result에서 정렬 필드 정보 추출
                sort_fields = result.get("sort_fields", [])

            except Exception as e:
                raise ValueError(f"정렬 적용 실패: {str(e)}")

            # 저장 처리
            saved = False
            if save:
                try:
                    book.save()
                    saved = True
                except Exception:
                    # 저장 실패해도 정렬은 적용된 상태
                    pass

            # 정렬 결과 정보 구성
            sort_info = {
                "table_name": table_name,
                "sheet": target_sheet.name,
                "sort_fields": [
                    {
                        "column": field.get("column", field.get("column_name", "")),
                        "order": field.get("order", "asc"),
                        "position": idx + 1,
                    }
                    for idx, field in enumerate(sort_fields)
                ],
                "total_sort_fields": len(sort_fields),
                "saved": saved,
            }

            # 워크북 정보 추가
            workbook_info = {
                "name": normalize_path(book.name),
                "full_name": normalize_path(book.fullname),
                "saved": getattr(book, "saved", True),
            }

            # 데이터 구성
            data_content = {
                "sort_result": sort_info,
                "workbook": workbook_info,
            }

            # 성공 메시지 생성
            sort_desc = ", ".join([f"{field['column']} ({field['order']})" for field in sort_fields])
            save_status = "저장됨" if saved else ("저장하지 않음" if not save else "저장 실패")
            message = f"테이블 '{table_name}'을 정렬했습니다: {sort_desc} ({save_status})"

            # 성공 응답 생성
            response = create_success_response(
                data=data_content,
                command="table-sort",
                message=message,
                execution_time_ms=timer.execution_time_ms,
                book=book,
            )

            # 출력 형식에 따른 결과 반환
            if output_format == "json":
                typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
            else:  # text 형식
                sort_result = sort_info
                wb = workbook_info

                typer.echo(f"✅ {message}")
                typer.echo()
                typer.echo(f"📁 워크북: {wb['name']}")
                typer.echo(f"📄 시트: {sort_result['sheet']}")
                typer.echo(f"🏷️ 테이블: {sort_result['table_name']}")
                typer.echo(f"📊 정렬 필드:")

                for field in sort_result["sort_fields"]:
                    order_emoji = "⬆️" if field["order"] == "asc" else "⬇️"
                    typer.echo(f"   {field['position']}. {field['column']} {order_emoji} {field['order']}")

                if saved:
                    typer.echo(f"💾 저장: ✅ 완료")
                elif not save:
                    typer.echo(f"💾 저장: ⚠️ 저장하지 않음 (--no-save 옵션)")
                else:
                    typer.echo(f"💾 저장: ❌ 실패")

    except FileNotFoundError as e:
        error_response = create_error_response(e, "table-sort")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ 파일을 찾을 수 없습니다", err=True)
        raise typer.Exit(1)

    except ValueError as e:
        error_response = create_error_response(e, "table-sort")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)

    except Exception as e:
        error_response = create_error_response(e, "table-sort")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ 예기치 않은 오류: {str(e)}", err=True)
            typer.echo(
                "💡 Excel이 설치되어 있는지 확인하고, 파일이 다른 프로그램에서 사용 중이지 않은지 확인하세요.", err=True
            )
        raise typer.Exit(1)

    finally:
        # 워크북 정리 - 활성 워크북이나 이름으로 접근한 경우 앱 종료하지 않음
        if book and not visible and file_path:
            try:
                book.app.quit()
            except:
                pass


if __name__ == "__main__":
    typer.run(table_sort)
