"""
Excel 테이블 목록 조회 명령어 (Typer 버전)
워크북의 모든 Excel Table(ListObject) 정보 조회
"""

import json
import platform
from pathlib import Path
from typing import Optional

import typer
import xlwings as xw

from pyhub_office_automation.version import get_version

from .utils import (
    ExecutionTimer,
    create_error_response,
    create_success_response,
    get_or_open_workbook,
    get_sheet,
    normalize_path,
)


def table_list(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="Excel 파일의 절대 경로"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help="열린 워크북 이름으로 접근"),
    sheet: Optional[str] = typer.Option(None, "--sheet", help="특정 시트만 조회 (미지정시 모든 시트)"),
    detailed: bool = typer.Option(False, "--detailed", help="상세 정보 포함 (범위, 스타일, 헤더 등)"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택"),
    visible: bool = typer.Option(False, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부"),
):
    """
    워크북의 모든 Excel Table(ListObject) 목록을 조회합니다.

    Excel Table 정보를 확인하여 피벗테이블 생성이나 데이터 분석 작업에 활용할 수 있습니다.
    Windows 전용 기능으로, macOS에서는 에러가 발생합니다.

    \b
    워크북 접근 방법:
      • 옵션 없음: 활성 워크북 자동 사용 (기본값)
      • --file-path: 파일 경로로 워크북 열기
      • --workbook-name: 열린 워크북 이름으로 접근

    \b
    출력 정보:
      • 기본: 테이블 이름, 시트명, 간단한 범위 정보
      • --detailed: 스타일, 헤더 여부, 행/열 수, 데이터 범위 등 상세 정보

    \b
    사용 예제:
      # 전체 워크북의 테이블 목록
      oa excel table-list

      # 상세 정보 포함
      oa excel table-list --detailed

      # 특정 시트만 조회
      oa excel table-list --sheet "Data" --detailed

      # 특정 파일의 테이블 목록
      oa excel table-list --file-path "sales.xlsx" --detailed

      # 특정 열린 워크북 조회
      oa excel table-list --workbook-name "Report.xlsx"
    """
    book = None
    try:
        with ExecutionTimer() as timer:
            # 플랫폼 확인
            if platform.system() != "Windows":
                raise ValueError("Excel Table 조회는 Windows에서만 지원됩니다.")

            # 워크북 연결
            book = get_or_open_workbook(file_path=file_path, workbook_name=workbook_name, visible=visible)

            # 조회할 시트 목록 결정
            if sheet:
                target_sheets = [get_sheet(book, sheet)]
            else:
                target_sheets = list(book.sheets)

            # 모든 테이블 정보 수집
            all_tables = []
            total_tables = 0

            for sheet_obj in target_sheets:
                sheet_tables = []

                try:
                    # 시트의 모든 테이블 조회
                    for table in sheet_obj.tables:
                        table_info = {
                            "name": table.name,
                            "sheet": sheet_obj.name,
                        }

                        # 상세 정보 추가
                        if detailed:
                            try:
                                # 기본 정보
                                table_info.update(
                                    {
                                        "range": table.range.address,
                                        "row_count": table.range.rows.count,
                                        "column_count": table.range.columns.count,
                                    }
                                )

                                # COM API를 통한 추가 정보 (Windows만)
                                try:
                                    list_object = None
                                    for lo in sheet_obj.api.ListObjects:
                                        if lo.Name == table.name:
                                            list_object = lo
                                            break

                                    if list_object:
                                        table_info.update(
                                            {
                                                "has_headers": list_object.HeaderRowRange is not None,
                                                "style": getattr(list_object, "TableStyle", "Unknown"),
                                                "data_range": (
                                                    list_object.DataBodyRange.Address if list_object.DataBodyRange else None
                                                ),
                                                "header_range": (
                                                    list_object.HeaderRowRange.Address if list_object.HeaderRowRange else None
                                                ),
                                                "total_range": (
                                                    list_object.TotalsRowRange.Address if list_object.TotalsRowRange else None
                                                ),
                                            }
                                        )
                                except:
                                    # COM API 접근 실패 시 기본값 설정
                                    table_info.update(
                                        {
                                            "has_headers": True,
                                            "style": "Unknown",
                                            "data_range": None,
                                            "header_range": None,
                                            "total_range": None,
                                        }
                                    )

                            except Exception as e:
                                # 상세 정보 수집 실패 시 기본 정보만 포함
                                table_info.update({"error": f"상세 정보 수집 실패: {str(e)}"})

                        sheet_tables.append(table_info)
                        total_tables += 1

                except Exception as e:
                    # 시트 접근 실패 시 에러 정보 추가
                    sheet_tables.append({"sheet": sheet_obj.name, "error": f"시트 접근 실패: {str(e)}"})

                if sheet_tables or not sheet:  # 특정 시트 지정했거나 테이블이 있는 경우만 추가
                    all_tables.extend(sheet_tables)

            # 워크북 정보
            workbook_info = {
                "name": normalize_path(book.name),
                "full_name": normalize_path(book.fullname),
                "saved": getattr(book, "saved", True),
                "sheet_count": len(book.sheets),
            }

            # 요약 정보
            summary = {
                "total_tables": total_tables,
                "sheets_with_tables": len(set(table.get("sheet") for table in all_tables if "error" not in table)),
                "sheets_scanned": len(target_sheets),
            }

            # 데이터 구성
            data_content = {
                "tables": all_tables,
                "summary": summary,
                "workbook": workbook_info,
                "query": {
                    "sheet_filter": sheet,
                    "detailed": detailed,
                },
            }

            # 성공 메시지 생성
            if sheet:
                message = f"시트 '{sheet}'에서 {total_tables}개의 Excel Table을 찾았습니다"
            else:
                message = f"워크북에서 총 {total_tables}개의 Excel Table을 찾았습니다"

            # 성공 응답 생성
            response = create_success_response(
                data=data_content,
                command="table-list",
                message=message,
                execution_time_ms=timer.execution_time_ms,
                book=book,
            )

            # 출력 형식에 따른 결과 반환
            if output_format == "json":
                typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
            else:  # text 형식
                wb = workbook_info
                sum_info = summary

                typer.echo(f"✅ {message}")
                typer.echo()
                typer.echo(f"📁 워크북: {wb['name']}")
                typer.echo(
                    f"📊 요약: {sum_info['total_tables']}개 테이블, {sum_info['sheets_with_tables']}/{sum_info['sheets_scanned']}개 시트"
                )

                if all_tables:
                    typer.echo()
                    typer.echo("📋 Excel Table 목록:")
                    typer.echo("-" * 50)

                    current_sheet = None
                    for table in all_tables:
                        if "error" in table:
                            typer.echo(f"❌ {table['sheet']}: {table['error']}")
                            continue

                        # 시트별 그룹핑
                        if table["sheet"] != current_sheet:
                            if current_sheet is not None:
                                typer.echo()
                            typer.echo(f"📄 {table['sheet']}:")
                            current_sheet = table["sheet"]

                        # 테이블 정보 출력
                        if detailed:
                            typer.echo(f"  🏷️ {table['name']}")
                            if "range" in table:
                                typer.echo(f"     📍 범위: {table['range']}")
                                typer.echo(f"     📊 크기: {table['row_count']}행 × {table['column_count']}열")
                            if "style" in table:
                                typer.echo(f"     🎨 스타일: {table['style']}")
                                typer.echo(f"     📋 헤더: {'있음' if table.get('has_headers', True) else '없음'}")
                            if "data_range" in table and table["data_range"]:
                                typer.echo(f"     📄 데이터: {table['data_range']}")
                        else:
                            typer.echo(f"  🏷️ {table['name']}")
                else:
                    typer.echo()
                    typer.echo("📋 Excel Table이 없습니다.")

    except FileNotFoundError as e:
        error_response = create_error_response(e, "table-list")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ 파일을 찾을 수 없습니다", err=True)
        raise typer.Exit(1)

    except ValueError as e:
        error_response = create_error_response(e, "table-list")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)

    except Exception as e:
        error_response = create_error_response(e, "table-list")
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
    typer.run(table_list)
