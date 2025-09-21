"""
Excel 워크북 열기 명령어 (Typer 버전)
AI 에이전트와의 연동을 위한 구조화된 출력 제공
"""

import json
import sys
from pathlib import Path
from typing import Optional

import typer
import xlwings as xw

from pyhub_office_automation.version import get_version

from .utils import ExecutionTimer, create_error_response, create_success_response, get_or_open_workbook, normalize_path


def workbook_open(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="열 Excel 파일의 절대 경로"),
    use_active: bool = typer.Option(False, "--use-active", help="현재 활성 워크북 정보를 가져옵니다"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help="열린 워크북 이름으로 찾기"),
    visible: bool = typer.Option(True, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택"),
):
    """
    Excel 워크북을 열거나 기존 워크북의 정보를 가져옵니다.

    다음 방법 중 하나를 사용할 수 있습니다:
    - --file-path: 지정된 경로의 파일을 엽니다
    - --use-active: 현재 활성 워크북의 정보를 가져옵니다
    - --workbook-name: 이미 열린 워크북을 이름으로 찾습니다

    예제:
        oa excel workbook-open --file-path "data.xlsx"
        oa excel workbook-open --use-active
        oa excel workbook-open --workbook-name "Sales.xlsx"
    """
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
            book = get_or_open_workbook(
                file_path=file_path, workbook_name=workbook_name, use_active=use_active, visible=visible
            )

            # 앱 객체 가져오기
            app = book.app

            # 시트 정보 수집
            sheets_info = []
            for sheet in book.sheets:
                try:
                    # 시트의 사용된 범위 정보
                    used_range = sheet.used_range
                    if used_range:
                        last_cell = used_range.last_cell.address
                        row_count = used_range.rows.count
                        col_count = used_range.columns.count
                    else:
                        last_cell = "A1"
                        row_count = 0
                        col_count = 0

                    sheet_info = {
                        "name": sheet.name,
                        "index": sheet.index,
                        "used_range": used_range.address if used_range else None,
                        "last_cell": last_cell,
                        "row_count": row_count,
                        "column_count": col_count,
                        "is_active": sheet == book.sheets.active,
                    }
                    sheets_info.append(sheet_info)

                except Exception as e:
                    # 개별 시트 정보 수집 실패 시 기본 정보만 포함
                    sheets_info.append(
                        {"name": sheet.name, "index": getattr(sheet, "index", -1), "error": f"시트 정보 수집 실패: {str(e)}"}
                    )

            # 워크북 정보 구성
            workbook_info = {
                "name": normalize_path(book.name),
                "full_name": normalize_path(book.fullname),
                "saved": getattr(book, "saved", True),
                "sheet_count": len(book.sheets),
                "active_sheet": book.sheets.active.name if book.sheets.active else None,
                "sheets": sheets_info,
            }

            # 파일 정보 추가 (실제 파일이 있는 경우)
            try:
                if hasattr(book, "fullname") and book.fullname:
                    file_path_info = Path(book.fullname)
                    if file_path_info.exists():
                        file_stat = file_path_info.stat()
                        workbook_info.update(
                            {
                                "file_size_bytes": file_stat.st_size,
                                "file_extension": file_path_info.suffix.lower(),
                                "is_read_only": not file_path_info.stat().st_mode & 0o200,
                            }
                        )
            except (OSError, AttributeError):
                # 새 워크북이거나 파일 정보 접근 불가능한 경우
                pass

            # 애플리케이션 정보
            app_info = {
                "version": getattr(app, "version", "Unknown"),
                "visible": getattr(app, "visible", visible),
                "calculation_mode": getattr(app, "calculation", "Unknown"),
            }

            # 데이터 구성
            data_content = {
                "workbook": workbook_info,
                "application": app_info,
                "connection_method": "file_path" if file_path else ("active" if use_active else "workbook_name"),
            }

            # 성공 메시지
            if use_active:
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
