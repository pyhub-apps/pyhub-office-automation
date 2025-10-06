"""
Excel 테이블 정렬 상태 조회 명령어 (Typer 버전)
Excel Table(ListObject)에 적용된 정렬 상태 확인
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


def table_sort_info(
    table_name: str = typer.Option(..., "--table-name", help="정렬 상태를 확인할 테이블 이름"),
    file_path: Optional[str] = typer.Option(None, "--file-path", help="Excel 파일의 절대 경로"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help="열린 워크북 이름으로 접근"),
    sheet: Optional[str] = typer.Option(None, "--sheet", help="시트 이름 (미지정시 테이블 이름으로 검색)"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택"),
    visible: bool = typer.Option(False, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부"),
):
    """
    Excel Table에 현재 적용된 정렬 상태를 조회합니다.

    적용된 정렬 필드, 정렬 순서, 정렬 우선순위 등의 정보를 확인할 수 있습니다.
    Windows 전용 기능으로, macOS에서는 에러가 발생합니다.

    \b
    워크북 접근 방법:
      • 옵션 없음: 활성 워크북 자동 사용 (기본값)
      • --file-path: 파일 경로로 워크북 열기
      • --workbook-name: 열린 워크북 이름으로 접근

    \b
    출력 정보:
      • 정렬 필드 목록 (컬럼명, 정렬 순서, 우선순위)
      • 정렬 적용 여부
      • 테이블 기본 정보
      • 정렬 없을 시 빈 배열 반환

    \b
    사용 예제:
      # 테이블 정렬 상태 확인
      oa excel table-sort-info --table-name "SalesData"

      # 특정 시트의 테이블 확인
      oa excel table-sort-info --table-name "ProductTable" --sheet "Products"

      # 특정 파일의 테이블 확인
      oa excel table-sort-info --table-name "DataTable" --file-path "report.xlsx"

      # 텍스트 형식으로 출력
      oa excel table-sort-info --table-name "SalesData" --format text
    """
    book = None
    try:
        with ExecutionTimer() as timer:
            # 플랫폼 확인
            if platform.system() != "Windows":
                raise ValueError("Excel Table 정렬 상태 조회는 Windows에서만 지원됩니다.")

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

            # 정렬 상태 조회 (Engine Layer 사용)
            sort_fields = []
            has_sort = False
            header_values = []

            try:
                result = engine.get_table_sort_info(workbook=book.api, sheet_name=target_sheet.name, table_name=table_name)

                sort_fields = result.get("sort_fields", [])
                has_sort = result.get("has_sort", False)
                header_values = result.get("headers", [])

            except Exception as e:
                # 정렬 정보 조회 실패 시 정렬 없음으로 처리
                has_sort = False
                sort_fields = []

            # 테이블 기본 정보
            table_info = {
                "name": table_name,
                "sheet": target_sheet.name,
                "range": target_table.range.address,
                "row_count": target_table.range.rows.count,
                "column_count": target_table.range.columns.count,
                "has_headers": len(header_values) > 0,
                "headers": header_values if header_values else [],
            }

            # 정렬 상태 정보
            sort_status = {
                "has_sort": has_sort,
                "sort_fields": sort_fields,
                "total_sort_fields": len(sort_fields),
                "sort_applied": has_sort and len(sort_fields) > 0,
            }

            # 워크북 정보 추가
            workbook_info = {
                "name": normalize_path(book.name),
                "full_name": normalize_path(book.fullname),
                "saved": getattr(book, "saved", True),
            }

            # 데이터 구성
            data_content = {
                "table": table_info,
                "sort_status": sort_status,
                "workbook": workbook_info,
            }

            # 성공 메시지 생성
            if has_sort and sort_fields:
                sort_desc = ", ".join([f"{field['column']} ({field['order']})" for field in sort_fields])
                message = f"테이블 '{table_name}'에 정렬이 적용되어 있습니다: {sort_desc}"
            else:
                message = f"테이블 '{table_name}'에 적용된 정렬이 없습니다"

            # 성공 응답 생성
            response = create_success_response(
                data=data_content,
                command="table-sort-info",
                message=message,
                execution_time_ms=timer.execution_time_ms,
                book=book,
            )

            # 출력 형식에 따른 결과 반환
            if output_format == "json":
                typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
            else:  # text 형식
                table = table_info
                sort_status_data = sort_status
                wb = workbook_info

                typer.echo(f"📊 {message}")
                typer.echo()
                typer.echo(f"📁 워크북: {wb['name']}")
                typer.echo(f"📄 시트: {table['sheet']}")
                typer.echo(f"🏷️ 테이블: {table['name']}")
                typer.echo(f"📍 범위: {table['range']}")
                typer.echo(f"📊 크기: {table['row_count']}행 × {table['column_count']}열")
                typer.echo()

                if sort_status_data["has_sort"] and sort_status_data["sort_fields"]:
                    typer.echo(f"🔀 정렬 상태: ✅ 적용됨 ({len(sort_status_data['sort_fields'])}개 필드)")
                    typer.echo(f"📋 정렬 필드:")
                    for field in sort_status_data["sort_fields"]:
                        order_emoji = "⬆️" if field["order"] == "asc" else "⬇️"
                        typer.echo(f"   {field['priority']}. {field['column']} {order_emoji} {field['order']}")
                else:
                    typer.echo(f"🔀 정렬 상태: ❌ 정렬 없음")

                typer.echo()
                if table["has_headers"] and table["headers"]:
                    typer.echo(f"📋 헤더: {', '.join(table['headers'])}")

    except FileNotFoundError as e:
        error_response = create_error_response(e, "table-sort-info")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ 파일을 찾을 수 없습니다", err=True)
        raise typer.Exit(1)

    except ValueError as e:
        error_response = create_error_response(e, "table-sort-info")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)

    except Exception as e:
        error_response = create_error_response(e, "table-sort-info")
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
        if book is not None and not visible and file_path:
            try:
                book.app.quit()
            except:
                pass


if __name__ == "__main__":
    typer.run(table_sort_info)
