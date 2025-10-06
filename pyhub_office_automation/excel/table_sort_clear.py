"""
Excel 테이블 정렬 해제 명령어 (Typer 버전)
Excel Table(ListObject)의 정렬 상태를 초기화하고 원래 순서로 복원
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


def table_sort_clear(
    table_name: str = typer.Option(..., "--table-name", help="정렬을 해제할 테이블 이름"),
    file_path: Optional[str] = typer.Option(None, "--file-path", help="Excel 파일의 절대 경로"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help="열린 워크북 이름으로 접근"),
    sheet: Optional[str] = typer.Option(None, "--sheet", help="시트 이름 (미지정시 테이블 이름으로 검색)"),
    save: bool = typer.Option(True, "--save/--no-save", help="정렬 해제 후 파일 저장 여부"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택"),
    visible: bool = typer.Option(False, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부"),
):
    """
    Excel Table의 정렬 상태를 해제하고 원래 순서로 복원합니다.

    적용된 모든 정렬을 제거하고 데이터를 원본 순서로 되돌립니다.
    Windows 전용 기능으로, macOS에서는 에러가 발생합니다.

    \b
    워크북 접근 방법:
      • 옵션 없음: 활성 워크북 자동 사용 (기본값)
      • --file-path: 파일 경로로 워크북 열기
      • --workbook-name: 열린 워크북 이름으로 접근

    \b
    정렬 해제 방법:
      • AutoFilter.Sort.SortFields.Clear() 메서드 사용
      • 모든 정렬 필드 제거
      • 데이터 원본 순서로 복원

    \b
    주의사항:
      • 정렬 해제 시 데이터가 입력된 원래 순서로 되돌아갑니다
      • 이미 정렬이 적용되지 않은 테이블에서는 변화가 없습니다
      • AutoFilter는 유지되며, 정렬 조건만 제거됩니다

    \b
    사용 예제:
      # 테이블 정렬 해제
      oa excel table-sort-clear --table-name "SalesData"

      # 특정 시트의 테이블 정렬 해제
      oa excel table-sort-clear --table-name "ProductTable" --sheet "Products"

      # 저장하지 않고 정렬만 해제
      oa excel table-sort-clear --table-name "TempData" --no-save

      # 텍스트 형식으로 출력
      oa excel table-sort-clear --table-name "DataTable" --format text
    """
    book = None
    try:
        with ExecutionTimer() as timer:
            # 플랫폼 확인
            if platform.system() != "Windows":
                raise ValueError("Excel Table 정렬 해제는 Windows에서만 지원됩니다.")

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

            # 정렬 해제 전 현재 정렬 상태 가져오기
            previous_sort_fields = []
            had_sort = False

            try:
                # 현재 정렬 상태 확인
                sort_info = engine.get_table_sort_info(workbook=book.api, sheet_name=target_sheet.name, table_name=table_name)

                previous_sort_fields = sort_info.get("sort_fields", [])
                had_sort = len(previous_sort_fields) > 0

            except Exception:
                # 정렬 상태 확인 실패 시 정렬 없음으로 처리
                had_sort = False

            # 정렬 해제 실행 (Engine Layer 사용)
            sort_cleared = False
            try:
                result = engine.clear_table_sort(workbook=book.api, sheet_name=target_sheet.name, table_name=table_name)

                sort_cleared = result.get("success", False)

            except Exception as e:
                raise ValueError(f"정렬 해제 실패: {str(e)}")

            # 저장 처리
            saved = False
            if save:
                try:
                    book.save()
                    saved = True
                except Exception:
                    # 저장 실패해도 정렬 해제는 완료된 상태
                    pass

            # 정렬 해제 결과 정보 구성
            clear_info = {
                "table_name": table_name,
                "sheet": target_sheet.name,
                "had_sort_before": had_sort,
                "previous_sort_fields": previous_sort_fields,
                "sort_cleared": sort_cleared,
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
                "clear_result": clear_info,
                "workbook": workbook_info,
            }

            # 성공 메시지 생성
            if had_sort:
                sort_desc = ", ".join([f"{field['column']} ({field['order']})" for field in previous_sort_fields])
                save_status = "저장됨" if saved else ("저장하지 않음" if not save else "저장 실패")
                message = f"테이블 '{table_name}'의 정렬을 해제했습니다. 이전 정렬: {sort_desc} ({save_status})"
            else:
                save_status = "저장됨" if saved else ("저장하지 않음" if not save else "저장 실패")
                message = f"테이블 '{table_name}'에 적용된 정렬이 없었습니다 ({save_status})"

            # 성공 응답 생성
            response = create_success_response(
                data=data_content,
                command="table-sort-clear",
                message=message,
                execution_time_ms=timer.execution_time_ms,
                book=book,
            )

            # 출력 형식에 따른 결과 반환
            if output_format == "json":
                typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
            else:  # text 형식
                clear_result = clear_info
                wb = workbook_info

                typer.echo(f"✅ {message}")
                typer.echo()
                typer.echo(f"📁 워크북: {wb['name']}")
                typer.echo(f"📄 시트: {clear_result['sheet']}")
                typer.echo(f"🏷️ 테이블: {clear_result['table_name']}")

                if clear_result["had_sort_before"]:
                    typer.echo(f"🔀 이전 정렬: ✅ 있었음")
                    if clear_result["previous_sort_fields"]:
                        typer.echo(f"📋 해제된 정렬 필드:")
                        for field in clear_result["previous_sort_fields"]:
                            order_emoji = "⬆️" if field["order"] == "asc" else "⬇️"
                            typer.echo(f"   {field['priority']}. {field['column']} {order_emoji} {field['order']}")
                else:
                    typer.echo(f"🔀 이전 정렬: ❌ 없었음")

                typer.echo(f"🧹 정렬 해제: {'✅ 완료' if clear_result['sort_cleared'] else '❌ 실패'}")

                if saved:
                    typer.echo(f"💾 저장: ✅ 완료")
                elif not save:
                    typer.echo(f"💾 저장: ⚠️ 저장하지 않음 (--no-save 옵션)")
                else:
                    typer.echo(f"💾 저장: ❌ 실패")

    except FileNotFoundError as e:
        error_response = create_error_response(e, "table-sort-clear")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ 파일을 찾을 수 없습니다", err=True)
        raise typer.Exit(1)

    except ValueError as e:
        error_response = create_error_response(e, "table-sort-clear")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)

    except Exception as e:
        error_response = create_error_response(e, "table-sort-clear")
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
    typer.run(table_sort_clear)
