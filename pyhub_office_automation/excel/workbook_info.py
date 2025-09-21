"""
Excel 워크북 상세 정보 조회 명령어 (Typer 버전)
특정 워크북의 상세 정보를 조회하여 AI 에이전트가 작업 컨텍스트를 파악할 수 있도록 지원
"""

import datetime
import json
import sys
from pathlib import Path
from typing import Optional

import typer
import xlwings as xw

from pyhub_office_automation.version import get_version

from .utils import ExecutionTimer, create_error_response, create_success_response, get_or_open_workbook, normalize_path


def workbook_info(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="조회할 Excel 파일의 절대 경로"),
    use_active: bool = typer.Option(False, "--use-active", help="현재 활성 워크북 정보를 조회합니다"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help="열린 워크북 이름으로 찾기"),
    include_sheets: bool = typer.Option(False, "--include-sheets", help="시트 목록 및 상세 정보 포함"),
    include_names: bool = typer.Option(False, "--include-names", help="정의된 이름(Named Ranges) 포함"),
    include_properties: bool = typer.Option(False, "--include-properties", help="파일 속성 정보 포함"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택"),
):
    """
    특정 Excel 워크북의 상세 정보를 조회합니다.

    다음 방법 중 하나를 사용할 수 있습니다:
    - --file-path: 지정된 경로의 파일 정보를 조회합니다 (파일을 열어야 함)
    - --use-active: 현재 활성 워크북의 정보를 조회합니다
    - --workbook-name: 이미 열린 워크북을 이름으로 찾아 조회합니다

    예제:
        oa excel workbook-info --use-active --include-sheets
        oa excel workbook-info --workbook-name "Sales.xlsx" --include-properties
        oa excel workbook-info --file-path "data.xlsx" --include-names
    """
    book = None
    try:
        # 옵션 검증
        options_count = sum([bool(file_path), use_active, bool(workbook_name)])
        if options_count == 0:
            raise ValueError("--file-path, --use-active, --workbook-name 중 하나는 반드시 지정해야 합니다")
        elif options_count > 1:
            raise ValueError("--file-path, --use-active, --workbook-name 중 하나만 지정할 수 있습니다")

        # 파일 경로가 지정된 경우 파일 검증
        if file_path:
            file_path_obj = Path(normalize_path(file_path)).resolve()
            if not file_path_obj.exists():
                raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path_obj}")
            if not file_path_obj.suffix.lower() in [".xlsx", ".xls", ".xlsm"]:
                raise ValueError(f"지원되지 않는 파일 형식입니다: {file_path_obj.suffix}")

        # 실행 시간 측정 시작
        with ExecutionTimer() as timer:
            # 워크북 가져오기
            book = get_or_open_workbook(file_path=file_path, workbook_name=workbook_name, use_active=use_active, visible=True)

            # 기본 워크북 정보 수집
            try:
                saved_status = book.saved
            except:
                saved_status = True  # 기본값으로 저장됨으로 가정

            try:
                app_visible = book.app.visible
            except:
                app_visible = True  # 기본값으로 보임으로 가정

            # 기본 워크북 정보
            workbook_data = {
                "name": normalize_path(book.name),
                "full_name": normalize_path(book.fullname),
                "saved": saved_status,
                "sheet_count": len(book.sheets),
                "active_sheet": book.sheets.active.name if book.sheets.active else None,
            }

            # 파일 속성 정보 추가
            if include_properties:
                try:
                    file_path_obj = Path(book.fullname)
                    if file_path_obj.exists():
                        file_stat = file_path_obj.stat()
                        workbook_data.update(
                            {
                                "file_properties": {
                                    "file_size_bytes": file_stat.st_size,
                                    "file_size_mb": round(file_stat.st_size / (1024 * 1024), 2),
                                    "last_modified": datetime.datetime.fromtimestamp(file_stat.st_mtime).isoformat(),
                                    "created": datetime.datetime.fromtimestamp(file_stat.st_ctime).isoformat(),
                                    "file_extension": file_path_obj.suffix.lower(),
                                    "is_read_only": not (file_stat.st_mode & 0o200),
                                }
                            }
                        )
                except (OSError, AttributeError) as e:
                    workbook_data["file_properties"] = {"error": f"파일 속성 수집 실패: {str(e)}"}

            # 시트 정보 추가
            if include_sheets:
                sheets_info = []
                for sheet in book.sheets:
                    try:
                        # 시트의 사용된 범위 정보
                        used_range = sheet.used_range
                        if used_range:
                            last_cell = used_range.last_cell.address
                            row_count = used_range.rows.count
                            col_count = used_range.columns.count
                            used_range_address = used_range.address
                        else:
                            last_cell = "A1"
                            row_count = 0
                            col_count = 0
                            used_range_address = None

                        # 테이블 정보 수집
                        tables_info = []
                        try:
                            for table in sheet.api.ListObjects:
                                tables_info.append(
                                    {
                                        "name": table.Name,
                                        "range": table.Range.Address,
                                        "header_row": table.HeaderRowRange.Address if table.HeaderRowRange else None,
                                    }
                                )
                        except:
                            pass  # 테이블이 없거나 접근 불가능한 경우

                        sheet_info = {
                            "name": sheet.name,
                            "index": sheet.index,
                            "is_active": sheet == book.sheets.active,
                            "used_range": used_range_address,
                            "last_cell": last_cell,
                            "row_count": row_count,
                            "column_count": col_count,
                            "is_visible": getattr(sheet, "visible", True),
                            "tables_count": len(tables_info),
                            "tables": tables_info if tables_info else [],
                        }

                        # 시트 색상 정보 (가능한 경우)
                        try:
                            if hasattr(sheet.api, "Tab") and hasattr(sheet.api.Tab, "Color"):
                                sheet_info["tab_color"] = sheet.api.Tab.Color
                        except:
                            pass

                        sheets_info.append(sheet_info)

                    except Exception as e:
                        sheets_info.append(
                            {
                                "name": getattr(sheet, "name", "Unknown"),
                                "index": getattr(sheet, "index", -1),
                                "error": f"시트 정보 수집 실패: {str(e)}",
                            }
                        )

                workbook_data["sheets"] = sheets_info

            # 정의된 이름(Named Ranges) 정보 추가
            if include_names:
                names_info = []
                try:
                    for name in book.names:
                        try:
                            name_info = {
                                "name": name.name,
                                "refers_to": name.refers_to,
                                "refers_to_range": name.refers_to_range.address if name.refers_to_range else None,
                                "is_visible": getattr(name, "visible", True),
                            }
                            names_info.append(name_info)
                        except Exception as e:
                            names_info.append(
                                {"name": getattr(name, "name", "Unknown"), "error": f"이름 정보 수집 실패: {str(e)}"}
                            )
                except Exception as e:
                    names_info = [{"error": f"정의된 이름 목록 수집 실패: {str(e)}"}]

                workbook_data["named_ranges"] = names_info
                workbook_data["named_ranges_count"] = len([n for n in names_info if "error" not in n])

            # 애플리케이션 정보
            app_info = {
                "version": getattr(book.app, "version", "Unknown"),
                "visible": app_visible,
                "calculation_mode": getattr(book.app, "calculation", "Unknown"),
            }

            # 데이터 구성
            data_content = {
                "workbook": workbook_data,
                "application": app_info,
                "connection_method": "file_path" if file_path else ("active" if use_active else "workbook_name"),
                "query_options": {
                    "include_sheets": include_sheets,
                    "include_names": include_names,
                    "include_properties": include_properties,
                },
            }

            # 성공 메시지
            detail_level = []
            if include_sheets:
                detail_level.append("시트 정보")
            if include_names:
                detail_level.append("정의된 이름")
            if include_properties:
                detail_level.append("파일 속성")

            if detail_level:
                detail_str = ", ".join(detail_level)
                message = f"워크북 '{workbook_data['name']}' 정보를 조회했습니다 (포함: {detail_str})"
            else:
                message = f"워크북 '{workbook_data['name']}' 기본 정보를 조회했습니다"

            # 성공 응답 생성
            response = create_success_response(
                data=data_content,
                command="workbook-info",
                message=message,
                execution_time_ms=timer.execution_time_ms,
                book=book,
            )

            # 출력 형식에 따른 결과 반환
            if output_format == "json":
                typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
            else:  # text 형식
                wb = workbook_data
                typer.echo(f"📊 {message}")
                typer.echo()
                typer.echo(f"📁 파일명: {wb['name']}")
                typer.echo(f"📍 경로: {wb['full_name']}")
                typer.echo(f"💾 저장 상태: {'저장됨' if wb['saved'] else '저장되지 않음'}")
                typer.echo(f"📄 시트 수: {wb['sheet_count']}")
                typer.echo(f"📑 활성 시트: {wb['active_sheet']}")

                if include_properties and "file_properties" in wb:
                    props = wb["file_properties"]
                    if "error" not in props:
                        typer.echo()
                        typer.echo("📋 파일 속성:")
                        typer.echo(f"  💽 크기: {props['file_size_mb']} MB ({props['file_size_bytes']} bytes)")
                        typer.echo(f"  📎 형식: {props['file_extension']}")
                        typer.echo(f"  🕐 수정: {props['last_modified']}")
                        typer.echo(f"  🔒 읽기전용: {'예' if props['is_read_only'] else '아니오'}")

                if include_names and "named_ranges" in wb:
                    typer.echo()
                    typer.echo(f"🏷️  정의된 이름: {wb.get('named_ranges_count', 0)}개")
                    for name in wb["named_ranges"]:
                        if "error" in name:
                            typer.echo(f"  ❌ {name['error']}")
                        else:
                            typer.echo(f"  • {name['name']} → {name['refers_to']}")

                if include_sheets and "sheets" in wb:
                    typer.echo()
                    typer.echo("📋 시트 상세 정보:")
                    for i, sheet in enumerate(wb["sheets"], 1):
                        if "error" in sheet:
                            typer.echo(f"  {i}. {sheet['name']} - ❌ {sheet['error']}")
                        else:
                            active_mark = " (활성)" if sheet["is_active"] else ""
                            typer.echo(f"  {i}. {sheet['name']}{active_mark}")
                            if sheet.get("used_range"):
                                typer.echo(
                                    f"     범위: {sheet['used_range']} ({sheet['row_count']}행 × {sheet['column_count']}열)"
                                )
                            if sheet.get("tables_count", 0) > 0:
                                typer.echo(f"     테이블: {sheet['tables_count']}개")

    except FileNotFoundError as e:
        error_response = create_error_response(e, "workbook-info")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ 파일을 찾을 수 없습니다: {file_path}", err=True)
        raise typer.Exit(1)

    except ValueError as e:
        error_response = create_error_response(e, "workbook-info")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)

    except Exception as e:
        error_response = create_error_response(e, "workbook-info")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ 예기치 않은 오류: {str(e)}", err=True)
            typer.echo("💡 Excel이 설치되어 있는지 확인하세요.", err=True)
        raise typer.Exit(1)

    finally:
        # 리소스 정리 - 파일을 직접 연 경우만 종료 고려
        if book and file_path and not use_active:
            try:
                book.app.quit()
            except:
                pass


if __name__ == "__main__":
    typer.run(workbook_info)
