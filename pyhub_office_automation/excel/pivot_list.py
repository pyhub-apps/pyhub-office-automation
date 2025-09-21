"""
피벗테이블 목록 조회 명령어
워크북 내 모든 피벗테이블의 정보를 조회
"""

import json
import platform
import sys
from pathlib import Path
from typing import Optional

import typer
import xlwings as xw

from pyhub_office_automation.version import get_version

from .utils import (
    create_error_response,
    create_success_response,
    format_output,
    get_or_open_workbook,
    get_sheet,
    get_workbook,
    normalize_path,
)


def pivot_list(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="조회할 Excel 파일의 절대 경로"),
    use_active: bool = typer.Option(False, "--use-active", help="현재 활성 워크북 사용"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help='열린 워크북 이름으로 접근 (예: "Sales.xlsx")'),
    sheet: Optional[str] = typer.Option(None, "--sheet", help="특정 시트의 피벗테이블만 조회 (지정하지 않으면 전체 워크북)"),
    include_details: bool = typer.Option(False, "--include-details", help="피벗테이블 상세 정보 포함 여부 (기본값: False)"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택 (json/text)"),
    visible: bool = typer.Option(False, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부 (기본값: False)"),
):
    """
    워크북 내 모든 피벗테이블의 목록과 정보를 조회합니다.

    Windows와 macOS 모두에서 작동하지만, 상세 정보는 Windows에서 더 많이 제공됩니다.

    워크북 접근 방법:
    - --file-path: 파일 경로로 워크북 열기
    - --use-active: 현재 활성 워크북 사용
    - --workbook-name: 열린 워크북 이름으로 접근

    예제:
        oa excel pivot-list --file-path "sales.xlsx"
        oa excel pivot-list --use-active --include-details
        oa excel pivot-list --workbook-name "Report.xlsx" --sheet "Dashboard"
    """
    book = None

    try:
        # 워크북 연결
        book = get_or_open_workbook(file_path=file_path, workbook_name=workbook_name, use_active=use_active, visible=visible)

        pivot_tables = []

        # 특정 시트 또는 전체 워크북 조회
        if sheet:
            # 특정 시트만 조회
            target_sheet = get_sheet(book, sheet)
            sheets_to_check = [target_sheet]
        else:
            # 전체 워크북 조회
            sheets_to_check = book.sheets

        # 각 시트에서 피벗테이블 찾기
        for ws in sheets_to_check:
            try:
                if platform.system() == "Windows":
                    # Windows에서는 COM API 사용
                    sheet_pivots = []
                    for pivot_table in ws.api.PivotTables():
                        pivot_info = {
                            "name": pivot_table.Name,
                            "sheet": ws.name,
                            "location": pivot_table.TableRange1.Address if hasattr(pivot_table, "TableRange1") else "Unknown",
                        }

                        if include_details:
                            try:
                                # 상세 정보 수집
                                pivot_info.update(
                                    {
                                        "source_data": (
                                            pivot_table.SourceData if hasattr(pivot_table, "SourceData") else "Unknown"
                                        ),
                                        "row_fields": (
                                            [field.Name for field in pivot_table.RowFields]
                                            if hasattr(pivot_table, "RowFields")
                                            else []
                                        ),
                                        "column_fields": (
                                            [field.Name for field in pivot_table.ColumnFields]
                                            if hasattr(pivot_table, "ColumnFields")
                                            else []
                                        ),
                                        "data_fields": (
                                            [field.Name for field in pivot_table.DataFields]
                                            if hasattr(pivot_table, "DataFields")
                                            else []
                                        ),
                                        "page_fields": (
                                            [field.Name for field in pivot_table.PageFields]
                                            if hasattr(pivot_table, "PageFields")
                                            else []
                                        ),
                                        "refresh_date": (
                                            str(pivot_table.RefreshDate) if hasattr(pivot_table, "RefreshDate") else None
                                        ),
                                        "cache_index": pivot_table.CacheIndex if hasattr(pivot_table, "CacheIndex") else None,
                                    }
                                )
                            except Exception as e:
                                pivot_info["details_error"] = f"상세 정보 수집 실패: {str(e)}"

                        sheet_pivots.append(pivot_info)

                    pivot_tables.extend(sheet_pivots)

                else:
                    # macOS에서는 제한적 정보만 제공
                    # xlwings를 통한 범위 스캔으로 피벗테이블 추정 (완벽하지 않음)
                    used_range = ws.used_range
                    if used_range:
                        # 간단한 휴리스틱: 사용된 범위 내에서 피벗테이블로 보이는 구조 찾기
                        pivot_info = {
                            "name": f"PivotTable_估算_{ws.name}",
                            "sheet": ws.name,
                            "location": "macOS에서는 정확한 위치 감지 불가",
                            "note": "macOS에서는 피벗테이블 정확한 감지가 제한적입니다",
                        }

                        if include_details:
                            pivot_info["limitation"] = "macOS에서는 상세 정보를 제공할 수 없습니다"

                        # 실제로 피벗테이블이 있는지 확실하지 않으므로 조건부 추가
                        # 여기서는 보수적으로 접근하여 빈 목록 반환
                        pass

            except Exception as e:
                # 시트별 오류는 로그만 남기고 계속 진행
                error_info = {
                    "sheet": ws.name,
                    "error": f"피벗테이블 조회 실패: {str(e)}",
                    "note": "이 시트에서는 피벗테이블을 찾을 수 없습니다",
                }
                pivot_tables.append(error_info)

        # 응답 데이터 구성
        data_content = {
            "pivot_tables": pivot_tables,
            "total_count": len([pt for pt in pivot_tables if "error" not in pt]),
            "error_count": len([pt for pt in pivot_tables if "error" in pt]),
            "scanned_sheets": [ws.name for ws in sheets_to_check],
            "platform": platform.system(),
            "details_included": include_details,
            "file_info": {
                "path": (
                    str(Path(normalize_path(file_path)).resolve())
                    if file_path
                    else (normalize_path(book.fullname) if hasattr(book, "fullname") else None)
                ),
                "name": Path(normalize_path(file_path)).name if file_path else normalize_path(book.name),
                "sheet_count": len(book.sheets),
            },
        }

        # macOS 제한사항 안내
        if platform.system() != "Windows":
            data_content["platform_limitation"] = (
                "macOS에서는 피벗테이블 정확한 감지가 제한적입니다. Windows 환경에서 더 정확한 정보를 확인할 수 있습니다."
            )

        # 성공 메시지 구성
        message = f"{data_content['total_count']}개의 피벗테이블을 찾았습니다"
        if data_content["error_count"] > 0:
            message += f" ({data_content['error_count']}개 시트에서 오류 발생)"

        response = create_success_response(data=data_content, command="pivot-list", message=message)

        # 출력 형식에 따른 결과 반환
        if output_format == "json":
            typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
        else:  # text 형식
            typer.echo(f"✅ 피벗테이블 목록 조회 완료")
            typer.echo(f"📄 파일: {data_content['file_info']['name']}")
            typer.echo(f"📊 총 피벗테이블: {data_content['total_count']}개")
            typer.echo(f"🔍 조회 시트: {', '.join(data_content['scanned_sheets'])}")

            if platform.system() != "Windows":
                typer.echo("⚠️ macOS에서는 피벗테이블 정확한 감지가 제한적입니다")

            typer.echo()

            if data_content["total_count"] > 0:
                for i, pivot in enumerate([pt for pt in pivot_tables if "error" not in pt], 1):
                    typer.echo(f"{i}. 📋 {pivot['name']}")
                    typer.echo(f"   📍 위치: {pivot['sheet']}!{pivot.get('location', 'Unknown')}")

                    if include_details and "row_fields" in pivot:
                        if pivot["row_fields"]:
                            typer.echo(f"   📊 행 필드: {', '.join(pivot['row_fields'])}")
                        if pivot["column_fields"]:
                            typer.echo(f"   📊 열 필드: {', '.join(pivot['column_fields'])}")
                        if pivot["data_fields"]:
                            typer.echo(f"   📊 값 필드: {', '.join(pivot['data_fields'])}")
                        if pivot.get("refresh_date"):
                            typer.echo(f"   🔄 마지막 새로고침: {pivot['refresh_date']}")

                    typer.echo()
            else:
                typer.echo("📭 피벗테이블이 없습니다")

            if data_content["error_count"] > 0:
                typer.echo("❌ 오류가 발생한 시트:")
                for error_pt in [pt for pt in pivot_tables if "error" in pt]:
                    typer.echo(f"   {error_pt['sheet']}: {error_pt['error']}")

    except ValueError as e:
        error_response = create_error_response(e, "pivot-list")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)

    except RuntimeError as e:
        error_response = create_error_response(e, "pivot-list")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
            typer.echo(
                "💡 Excel이 설치되어 있는지 확인하고, 파일이 다른 프로그램에서 사용 중이지 않은지 확인하세요.", err=True
            )
        raise typer.Exit(1)

    except Exception as e:
        error_response = create_error_response(e, "pivot-list")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ 예기치 않은 오류: {str(e)}", err=True)
        raise typer.Exit(1)

    finally:
        # 워크북 정리 - 활성 워크북이나 이름으로 접근한 경우 앱 종료하지 않음
        if book and not visible and file_path:
            try:
                book.app.quit()
            except:
                pass
