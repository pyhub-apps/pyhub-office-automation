"""
Excel 워크북 열기 명령어 (Typer 버전)
AI 에이전트와의 연동을 위한 구조화된 출력 제공
"""

import json
import platform
import sys
from pathlib import Path
from typing import Optional

import typer

from pyhub_office_automation.version import get_version

from .engines import get_engine
from .utils import ExecutionTimer, create_error_response, create_success_response, normalize_path


def workbook_open(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="열 Excel 파일의 절대 경로"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help="열린 워크북 이름으로 찾기"),
    visible: bool = typer.Option(True, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택"),
):
    """
    Excel 워크북을 열거나 기존 워크북의 정보를 가져옵니다.

    워크북 접근 방법:
    - 옵션 없음: 활성 워크북 자동 사용 (기본값)
    - --file-path: 파일 경로로 워크북 열기
    - --workbook-name: 열린 워크북 이름으로 접근

    예제:
        oa excel workbook-open --file-path "data.xlsx"
        oa excel workbook-open
        oa excel workbook-open --workbook-name "Sales.xlsx"
    """
    try:
        # 옵션 검증 (이제 빈 옵션은 자동으로 활성 워크북 사용)
        options_count = sum([bool(file_path), bool(workbook_name)])
        if options_count > 1:
            raise ValueError("--file-path, --workbook-name 중 하나만 지정할 수 있습니다")

        # 파일 경로가 지정된 경우 파일 검증
        if file_path:
            file_path_obj = Path(normalize_path(file_path)).resolve()
            if not file_path_obj.exists():
                raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path_obj}")
            if not file_path_obj.suffix.lower() in [".xlsx", ".xls", ".xlsm"]:
                raise ValueError(f"지원되지 않는 파일 형식입니다: {file_path_obj.suffix}")

        # 실행 시간 측정 시작
        with ExecutionTimer() as timer:
            # Engine 획득
            engine = get_engine()

            # 워크북 가져오기
            if file_path:
                book = engine.open_workbook(file_path, visible=visible)
            elif workbook_name:
                book = engine.get_workbook_by_name(workbook_name)
            else:
                book = engine.get_active_workbook()

            # 기본 워크북 정보 가져오기 (Engine 사용)
            wb_info = engine.get_workbook_info(book)

            # 시트 정보 수집 (Windows COM을 통해)
            sheets_info = []
            if platform.system() == "Windows":
                try:
                    active_sheet_name = wb_info["active_sheet"]
                    for i, sheet_name in enumerate(wb_info["sheets"], start=1):
                        try:
                            ws = book.Sheets(sheet_name)
                            # 시트의 사용된 범위 정보
                            try:
                                used_range = ws.UsedRange
                                if used_range:
                                    last_cell = used_range.Cells(used_range.Rows.Count, used_range.Columns.Count).Address
                                    row_count = used_range.Rows.Count
                                    col_count = used_range.Columns.Count
                                    used_range_address = used_range.Address
                                else:
                                    last_cell = "A1"
                                    row_count = 0
                                    col_count = 0
                                    used_range_address = None
                            except:
                                last_cell = "A1"
                                row_count = 0
                                col_count = 0
                                used_range_address = None

                            sheet_info = {
                                "name": sheet_name,
                                "index": i,
                                "used_range": used_range_address,
                                "last_cell": last_cell,
                                "row_count": row_count,
                                "column_count": col_count,
                                "is_active": sheet_name == active_sheet_name,
                            }
                            sheets_info.append(sheet_info)

                        except Exception as e:
                            # 개별 시트 정보 수집 실패 시 기본 정보만 포함
                            sheets_info.append(
                                {
                                    "name": sheet_name,
                                    "index": i,
                                    "error": f"시트 정보 수집 실패: {str(e)}",
                                }
                            )
                except Exception as e:
                    # Windows에서 전체 시트 정보 수집 실패 시 기본 정보만 사용
                    for i, sheet_name in enumerate(wb_info["sheets"], start=1):
                        sheets_info.append(
                            {
                                "name": sheet_name,
                                "index": i,
                                "is_active": sheet_name == wb_info["active_sheet"],
                            }
                        )
            else:
                # macOS: 기본 정보만 사용
                active_sheet_name = wb_info["active_sheet"]
                for i, sheet_name in enumerate(wb_info["sheets"], start=1):
                    sheets_info.append(
                        {
                            "name": sheet_name,
                            "index": i,
                            "is_active": sheet_name == active_sheet_name,
                        }
                    )

            # 워크북 정보 구성
            workbook_info = {
                "name": normalize_path(wb_info["name"]),
                "full_name": normalize_path(wb_info["full_name"]),
                "saved": wb_info["saved"],
                "sheet_count": wb_info["sheet_count"],
                "active_sheet": wb_info["active_sheet"],
                "sheets": sheets_info,
            }

            # 파일 정보 추가 (실제 파일이 있는 경우)
            if wb_info.get("file_size_bytes"):
                workbook_info["file_size_bytes"] = wb_info["file_size_bytes"]

            if wb_info.get("full_name"):
                file_path_info = Path(wb_info["full_name"])
                if file_path_info.exists():
                    workbook_info["file_extension"] = file_path_info.suffix.lower()
                    workbook_info["is_read_only"] = not (file_path_info.stat().st_mode & 0o200)

            # 애플리케이션 정보 (Windows만)
            app_info = {}
            if platform.system() == "Windows":
                try:
                    app_info = {
                        "version": str(book.Application.Version),
                        "visible": bool(book.Application.Visible),
                        "calculation_mode": str(book.Application.Calculation),
                    }
                except:
                    app_info = {
                        "version": "Unknown",
                        "visible": visible,
                        "calculation_mode": "Unknown",
                    }
            else:
                # macOS
                app_info = {
                    "version": "Unknown",
                    "visible": visible,
                    "calculation_mode": "Unknown",
                }

            # 데이터 구성
            data_content = {
                "workbook": workbook_info,
                "application": app_info,
                "connection_method": "file_path" if file_path else ("workbook_name" if workbook_name else "active"),
            }

            # 성공 메시지
            if not file_path and not workbook_name:
                message = f"활성 워크북 '{workbook_info['name']}' 정보를 가져왔습니다"
            elif workbook_name:
                message = f"워크북 '{workbook_info['name']}' 정보를 가져왔습니다"
            else:
                message = f"워크북 '{workbook_info['name']}'을(를) 열었습니다"

            # 성공 응답 생성
            response = create_success_response(
                data=data_content,
                command="workbook-open",
                message=message,
                execution_time_ms=timer.execution_time_ms,
                book=book,
            )

            # 출력 형식에 따른 결과 반환
            if output_format == "json":
                typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
            else:  # text 형식
                wb = workbook_info
                typer.echo(f"📊 {message}")
                typer.echo()
                typer.echo(f"📁 파일명: {wb['name']}")
                typer.echo(f"📍 경로: {wb['full_name']}")
                typer.echo(f"💾 저장 상태: {'저장됨' if wb['saved'] else '저장되지 않음'}")
                typer.echo(f"📄 시트 수: {wb['sheet_count']}")
                typer.echo(f"📑 활성 시트: {wb['active_sheet']}")

                if "file_size_bytes" in wb:
                    size_mb = wb["file_size_bytes"] / (1024 * 1024)
                    typer.echo(f"💽 파일 크기: {size_mb:.1f} MB")
                    typer.echo(f"📎 파일 형식: {wb['file_extension']}")

                typer.echo()
                typer.echo("📋 시트 목록:")
                for i, sheet in enumerate(wb["sheets"], 1):
                    active_mark = " (활성)" if sheet.get("is_active") else ""
                    if "error" in sheet:
                        typer.echo(f"  {i}. {sheet['name']}{active_mark} - ❌ {sheet['error']}")
                    else:
                        typer.echo(f"  {i}. {sheet['name']}{active_mark}")
                        if sheet.get("used_range"):
                            typer.echo(
                                f"     범위: {sheet['used_range']} ({sheet['row_count']}행 × {sheet['column_count']}열)"
                            )

    except FileNotFoundError as e:
        error_response = create_error_response(e, "workbook-open")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ 파일을 찾을 수 없습니다: {file_path}", err=True)
        raise typer.Exit(1)

    except ValueError as e:
        error_response = create_error_response(e, "workbook-open")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)

    except Exception as e:
        error_response = create_error_response(e, "workbook-open")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ 예기치 않은 오류: {str(e)}", err=True)
            typer.echo("💡 Excel이 설치되어 있는지 확인하세요.", err=True)
        raise typer.Exit(1)


if __name__ == "__main__":
    typer.run(workbook_open)
