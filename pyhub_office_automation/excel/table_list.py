"""
Excel 테이블 목록 조회 명령어 (Typer 버전)
워크북의 모든 Excel Table(ListObject) 정보 조회
"""

import json
import platform
from pathlib import Path
from typing import Optional

import typer

from pyhub_office_automation.version import get_version

from .engines import get_engine
from .utils import ExecutionTimer, create_error_response, create_success_response


def truncate_sample_data(sample_data, max_length=50):
    """
    샘플 데이터의 각 셀 길이를 제한합니다.

    Args:
        sample_data: 샘플 데이터 리스트
        max_length: 최대 문자 길이

    Returns:
        list: 길이 제한된 샘플 데이터
    """
    if not sample_data:
        return []

    def truncate_cell_value(value):
        if value is None:
            return None
        str_value = str(value)
        return str_value[:max_length] + "..." if len(str_value) > max_length else str_value

    truncated_data = []
    for row in sample_data:
        if isinstance(row, list):
            truncated_row = [truncate_cell_value(cell) for cell in row]
        else:
            truncated_row = [truncate_cell_value(row)]
        truncated_data.append(truncated_row)

    return truncated_data


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

            # Engine 획득
            engine = get_engine()

            # 워크북 연결
            if file_path:
                book = engine.open_workbook(file_path, visible=visible)
            elif workbook_name:
                book = engine.get_workbook_by_name(workbook_name)
            else:
                book = engine.get_active_workbook()

            # 워크북 정보 조회
            wb_info = engine.get_workbook_info(book)

            # 테이블 목록 조회 (Engine 메서드 사용)
            table_infos = engine.list_tables(book, sheet=sheet)

            # 모든 테이블 정보 수집
            all_tables = []
            total_tables = len(table_infos)

            for table_info in table_infos:
                # TableInfo를 딕셔너리로 변환하고 추가 정보 포함
                table_dict = {
                    "name": table_info.name,
                    "sheet": table_info.sheet_name,
                    "range": table_info.address,
                    "row_count": table_info.row_count,
                    "column_count": table_info.column_count,
                    "has_headers": len(table_info.headers) > 0,
                    "data_rows": table_info.row_count - (1 if len(table_info.headers) > 0 else 0),
                    "columns": table_info.headers,
                    "sample_data": truncate_sample_data(table_info.sample_data) if table_info.sample_data else [],
                }

                # --detailed 옵션: Windows COM API로 추가 정보 조회
                if detailed:
                    try:
                        # COM API를 통한 상세 정보 (Windows만)
                        ws = book.Sheets(table_info.sheet_name)
                        list_object = ws.ListObjects(table_info.name)

                        # 스타일 정보
                        try:
                            style_name = (
                                list_object.TableStyle.Name
                                if hasattr(list_object.TableStyle, "Name")
                                else str(list_object.TableStyle)
                            )
                            table_dict["style"] = style_name
                        except:
                            table_dict["style"] = "Unknown"

                        # 상세 범위 정보
                        table_dict.update(
                            {
                                "data_range": list_object.DataBodyRange.Address if list_object.DataBodyRange else None,
                                "header_range": list_object.HeaderRowRange.Address if list_object.HeaderRowRange else None,
                                "total_range": list_object.TotalsRowRange.Address if list_object.TotalsRowRange else None,
                            }
                        )
                    except Exception as e:
                        table_dict["detailed_error"] = f"고급 정보 수집 실패: {str(e)}"

                all_tables.append(table_dict)

            # 워크북 정보
            workbook_info = {
                "name": wb_info["workbook"]["name"],
                "full_name": wb_info["workbook"]["full_name"],
                "saved": wb_info["workbook"]["saved"],
                "sheet_count": wb_info["workbook"]["sheet_count"],
            }

            # 요약 정보
            summary = {
                "total_tables": total_tables,
                "sheets_with_tables": len(set(t["sheet"] for t in all_tables)),
                "sheets_scanned": 1 if sheet else wb_info["workbook"]["sheet_count"],
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

                        # 테이블 정보 출력 (유용한 기본 정보를 모두 표시)
                        typer.echo(f"  🏷️ {table['name']}")

                        # 범위 정보 (항상 표시)
                        if "range" in table and table["range"] != "Unknown":
                            typer.echo(f"     📍 범위: {table['range']}")

                        # 크기 정보 (전체/데이터 구분하여 표시)
                        if table.get("row_count", 0) > 0 or table.get("column_count", 0) > 0:
                            total_rows = table["row_count"]
                            data_rows = table.get("data_rows", total_rows - 1)
                            columns = table["column_count"]
                            typer.echo(f"     📊 크기: {total_rows}행({data_rows}개 데이터) × {columns}열")

                        # 헤더 및 스타일 정보 (기본으로 표시)
                        if "has_headers" in table:
                            header_status = "있음" if table["has_headers"] else "없음"
                            typer.echo(f"     📋 헤더: {header_status}")

                        if "style" in table and table["style"] != "Unknown":
                            typer.echo(f"     🎨 스타일: {table['style']}")

                        # 컬럼 정보 (항상 표시)
                        if "columns" in table and table["columns"]:
                            columns_text = ", ".join(table["columns"])
                            typer.echo(f"     📋 컬럼 ({len(table['columns'])}개):")
                            typer.echo(f"       {columns_text}")

                        # 샘플 데이터 (항상 표시)
                        if "sample_data" in table and table["sample_data"]:
                            typer.echo(f"     📄 샘플 데이터 (상위 {len(table['sample_data'])}행):")
                            for i, row in enumerate(table["sample_data"], 1):
                                row_text = str(row)
                                typer.echo(f"       {i}. {row_text}")

                        # --detailed 옵션: 고급 범위 세부 정보만 추가 표시
                        if detailed:
                            if "data_range" in table and table["data_range"]:
                                typer.echo(f"     📄 데이터 범위: {table['data_range']}")
                            if "header_range" in table and table["header_range"]:
                                typer.echo(f"     📋 헤더 범위: {table['header_range']}")
                            if "total_range" in table and table["total_range"]:
                                typer.echo(f"     🔢 합계 범위: {table['total_range']}")
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
        # 워크북 정리 - 파일 경로로 열었고 visible=False인 경우에만 앱 종료
        if book and not visible and file_path:
            try:
                book.Application.Quit()
            except:
                pass


if __name__ == "__main__":
    typer.run(table_list)
