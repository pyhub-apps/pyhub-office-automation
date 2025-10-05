"""
Excel 워크시트 활성화 명령어 (Typer 버전)
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
from .utils import ExecutionTimer, create_error_response, create_success_response


def sheet_activate(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="열 Excel 파일의 절대 경로"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help="열린 워크북 이름으로 접근"),
    sheet: Optional[str] = typer.Option(None, "--sheet", help="활성화할 시트의 이름"),
    name: Optional[str] = typer.Option(None, "--name", help="[별칭] 활성화할 시트의 이름 (--sheet 사용 권장)"),
    index: Optional[int] = typer.Option(None, "--index", help="활성화할 시트의 인덱스 (0부터 시작)"),
    visible: bool = typer.Option(True, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택"),
):
    """
    Excel 워크북의 특정 시트를 활성화합니다.

    시트를 이름 또는 인덱스로 지정할 수 있습니다.
    활성화된 시트는 사용자에게 표시되는 현재 시트가 됩니다.

    워크북 접근 방법:
    - --file-path: 파일 경로로 워크북 열기
    - --workbook-name: 열린 워크북 이름으로 접근

    예제:
        oa excel sheet-activate --sheet "Sheet2"
        oa excel sheet-activate --file-path "data.xlsx" --index 1
        oa excel sheet-activate --workbook-name "Sales.xlsx" --sheet "Summary"
    """
    book = None
    try:
        # 옵션 우선순위 처리 (새 옵션 우선)
        sheet_name = sheet or name

        # 옵션 검증
        if sheet_name and index is not None:
            raise ValueError("--sheet(또는 --name)과 --index 옵션 중 하나만 지정할 수 있습니다")

        if not sheet_name and index is None:
            raise ValueError("--sheet(또는 --name) 또는 --index 중 하나는 반드시 지정해야 합니다")

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

            # 기본 워크북 정보 가져오기
            wb_info = engine.get_workbook_info(book)

            # 시트 목록에서 대상 시트 결정
            all_sheets_names = wb_info["sheets"]

            if sheet_name:
                # 이름으로 찾기
                if sheet_name not in all_sheets_names:
                    raise ValueError(f"시트 '{sheet_name}'을 찾을 수 없습니다. 사용 가능한 시트: {all_sheets_names}")
                target_sheet_name = sheet_name
            else:
                # 인덱스로 찾기 (0-based → 1-based 변환 필요)
                if index < 0 or index >= len(all_sheets_names):
                    raise ValueError(f"인덱스 {index}가 범위를 벗어났습니다. 사용 가능한 인덱스: 0-{len(all_sheets_names)-1}")
                target_sheet_name = all_sheets_names[index]

            # 활성화 전 정보 저장
            old_active_sheet_name = wb_info["active_sheet"]
            old_active_info = {
                "name": old_active_sheet_name,
                "index": all_sheets_names.index(old_active_sheet_name) if old_active_sheet_name in all_sheets_names else 0,
            }

            # Engine을 통해 시트 활성화
            engine.activate_sheet(book, target_sheet_name)

            # 활성화 후 워크북 정보 다시 가져오기
            wb_info_after = engine.get_workbook_info(book)
            new_active_sheet_name = wb_info_after["active_sheet"]

            # 시트 목록 구성 (활성 상태 표시)
            all_sheets = []
            for idx, sheet_nm in enumerate(all_sheets_names):
                all_sheets.append({"name": sheet_nm, "index": idx, "is_active": sheet_nm == new_active_sheet_name})

            # 활성화된 시트 정보
            activated_sheet_info = {
                "name": target_sheet_name,
                "index": all_sheets_names.index(target_sheet_name),
                "is_visible": True,  # 기본값
                "used_range": None,  # Engine이 제공하지 않으면 None
            }

            # 워크북 정보
            workbook_info = {
                "name": wb_info["name"],
                "full_name": wb_info["full_name"],
                "total_sheets": wb_info["sheet_count"],
            }

            new_active_info = {
                "name": new_active_sheet_name,
                "index": all_sheets_names.index(new_active_sheet_name) if new_active_sheet_name in all_sheets_names else 0,
            }

            # 데이터 구성
            data_content = {
                "activated_sheet": activated_sheet_info,
                "previous_active": old_active_info,
                "workbook": workbook_info,
                "all_sheets": all_sheets,
            }

            # 성공 메시지
            if sheet_name:
                message = f"시트 '{target_sheet_name}'을(를) 활성화했습니다"
            else:
                message = f"인덱스 {index}번 시트 '{target_sheet_name}'을(를) 활성화했습니다"

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
        # Engine이 리소스 관리를 담당하므로 추가 정리 불필요
        pass


if __name__ == "__main__":
    typer.run(sheet_activate)
