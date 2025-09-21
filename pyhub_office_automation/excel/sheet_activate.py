"""
Excel 워크시트 활성화 명령어 (Typer 버전)
AI 에이전트와의 연동을 위한 구조화된 출력 제공
"""

import json
import sys
from pathlib import Path
from typing import Optional

import typer
import xlwings as xw

from pyhub_office_automation.version import get_version

from .utils import ExecutionTimer, create_error_response, create_success_response, get_or_open_workbook, get_workbook


def sheet_activate(
    workbook: Optional[str] = typer.Option(None, "--workbook", help="워크북 파일 경로"),
    use_active: bool = typer.Option(False, "--use-active", help="현재 활성 워크북 사용"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help="열린 워크북 이름으로 접근"),
    name: Optional[str] = typer.Option(None, "--name", help="활성화할 시트의 이름"),
    index: Optional[int] = typer.Option(None, "--index", help="활성화할 시트의 인덱스 (1부터 시작)"),
    visible: bool = typer.Option(True, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택"),
):
    """
    Excel 워크북의 특정 시트를 활성화합니다.

    시트를 이름 또는 인덱스로 지정할 수 있습니다.
    활성화된 시트는 사용자에게 표시되는 현재 시트가 됩니다.

    워크북 접근 방법:
    - --workbook: 파일 경로로 워크북 열기 (기존 방식)
    - --use-active: 현재 활성 워크북 사용
    - --workbook-name: 열린 워크북 이름으로 접근

    예제:
        oa excel sheet-activate --use-active --name "Sheet2"
        oa excel sheet-activate --workbook "data.xlsx" --index 2
        oa excel sheet-activate --workbook-name "Sales.xlsx" --name "Summary"
    """
    book = None
    try:
        # 옵션 검증
        if name and index is not None:
            raise ValueError("--name과 --index 옵션 중 하나만 지정할 수 있습니다")

        if not name and index is None:
            raise ValueError("--name 또는 --index 중 하나는 반드시 지정해야 합니다")

        # 실행 시간 측정 시작
        with ExecutionTimer() as timer:
            # 워크북 연결
            book = get_or_open_workbook(
                file_path=workbook, workbook_name=workbook_name, use_active=use_active, visible=visible
            )

            # 기존 활성 시트 정보 수집
            old_active_sheet = book.sheets.active
            old_active_info = {"name": old_active_sheet.name, "index": old_active_sheet.index}

            # 시트 목록 수집
            all_sheets = []
            for sheet in book.sheets:
                all_sheets.append({"name": sheet.name, "index": sheet.index, "is_active": sheet == old_active_sheet})

            # 대상 시트 찾기 및 활성화
            target_sheet = None

            if name:
                # 이름으로 찾기
                try:
                    target_sheet = book.sheets[name]
                except KeyError:
                    available_names = [sheet.name for sheet in book.sheets]
                    raise ValueError(f"시트 '{name}'을 찾을 수 없습니다. 사용 가능한 시트: {available_names}")
            else:
                # 인덱스로 찾기
                try:
                    # xlwings는 1부터 시작하는 인덱스 사용
                    target_sheet = book.sheets[index]
                except IndexError:
                    sheet_count = len(book.sheets)
                    raise ValueError(f"인덱스 {index}가 범위를 벗어났습니다. 사용 가능한 인덱스: 1-{sheet_count}")

            # 시트 활성화
            target_sheet.activate()

            # 활성화 후 정보 수집
            new_active_sheet = book.sheets.active
            new_active_info = {"name": new_active_sheet.name, "index": new_active_sheet.index}

            # 시트 정보 업데이트
            for sheet_info in all_sheets:
                sheet_info["is_active"] = sheet_info["name"] == new_active_sheet.name

            # 활성화된 시트의 추가 정보
            activated_sheet_info = {
                "name": target_sheet.name,
                "index": target_sheet.index,
                "is_visible": getattr(target_sheet, "visible", True),
            }

            # 사용된 범위 정보 추가 (가능한 경우)
            try:
                used_range = target_sheet.used_range
                if used_range:
                    activated_sheet_info["used_range"] = {
                        "address": used_range.address,
                        "last_cell": used_range.last_cell.address,
                        "row_count": used_range.rows.count,
                        "column_count": used_range.columns.count,
                    }
                else:
                    activated_sheet_info["used_range"] = None
            except:
                activated_sheet_info["used_range"] = None

            # 워크북 정보
            workbook_info = {"name": book.name, "full_name": book.fullname, "total_sheets": len(book.sheets)}

            # 데이터 구성
            data_content = {
                "activated_sheet": activated_sheet_info,
                "previous_active": old_active_info,
                "workbook": workbook_info,
                "all_sheets": all_sheets,
            }

            # 성공 메시지
            if name:
                message = f"시트 '{target_sheet.name}'을(를) 활성화했습니다"
            else:
                message = f"인덱스 {index}번 시트 '{target_sheet.name}'을(를) 활성화했습니다"

            # 성공 응답 생성
            response = create_success_response(
                data=data_content,
                command="sheet-activate",
                message=message,
                execution_time_ms=timer.execution_time_ms,
                book=book,
            )

            # 출력 형식에 따른 결과 반환
            if output_format == "json":
                typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
            else:  # text 형식
                activated = activated_sheet_info
                wb = workbook_info

                typer.echo(f"✅ {message}")
                typer.echo()
                typer.echo(f"📁 워크북: {wb['name']}")
                typer.echo(f"📄 활성 시트: {activated['name']} (인덱스: {activated['index']})")

                if activated.get("used_range"):
                    used = activated["used_range"]
                    typer.echo(f"📊 사용된 범위: {used['address']} ({used['row_count']}행 × {used['column_count']}열)")
                else:
                    typer.echo(f"📊 사용된 범위: 없음 (빈 시트)")

                typer.echo()
                typer.echo(f"📋 전체 시트 목록 ({wb['total_sheets']}개):")
                for i, sheet in enumerate(all_sheets, 1):
                    active_mark = " ← 현재 활성" if sheet["is_active"] else ""
                    typer.echo(f"  {i}. {sheet['name']}{active_mark}")

    except ValueError as e:
        error_response = create_error_response(e, "sheet-activate")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)

    except Exception as e:
        error_response = create_error_response(e, "sheet-activate")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ 예기치 않은 오류: {str(e)}", err=True)
            typer.echo("💡 Excel이 설치되어 있는지 확인하고, 워크북이 열려있는지 확인하세요.", err=True)
        raise typer.Exit(1)

    finally:
        # 워크북 정리 - 활성 워크북이나 이름으로 접근한 경우 앱 종료하지 않음
        if book and not visible and workbook:
            try:
                book.app.quit()
            except:
                pass


if __name__ == "__main__":
    typer.run(sheet_activate)
