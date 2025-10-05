"""
Excel 워크북 목록 조회 명령어 (Engine 기반)
현재 열려있는 모든 워크북들의 목록과 기본 정보 제공
"""

import json
from typing import Optional

import typer

from pyhub_office_automation.version import get_version

from .engines import get_engine
from .utils import ExecutionTimer, create_error_response, create_success_response


def workbook_list(
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택"),
):
    """
    현재 열려있는 모든 Excel 워크북의 목록과 상세 정보를 조회합니다.

    각 워크북의 이름, 저장 상태, 파일 경로, 시트 수, 활성 시트 등의 정보를 제공합니다.

    예제:
        oa excel workbook-list
        oa excel workbook-list --format text
    """
    try:
        # 실행 시간 측정 시작
        with ExecutionTimer() as timer:
            # Engine 획득 (플랫폼 자동 감지)
            engine = get_engine()

            # 워크북 목록 조회
            workbooks = engine.get_workbooks()

            # 데이터 변환 (WorkbookInfo → dict)
            workbooks_data = []
            has_unsaved = False

            for wb_info in workbooks:
                workbook_dict = {
                    "name": wb_info.name,
                    "saved": wb_info.saved,
                    "full_name": wb_info.full_name,
                    "sheet_count": wb_info.sheet_count,
                    "active_sheet": wb_info.active_sheet,
                }

                # 선택적 정보 추가
                if wb_info.file_size_bytes is not None:
                    workbook_dict["file_size_bytes"] = wb_info.file_size_bytes

                if wb_info.last_modified is not None:
                    workbook_dict["last_modified"] = wb_info.last_modified

                if not wb_info.saved:
                    has_unsaved = True

                workbooks_data.append(workbook_dict)

            # 메시지 생성
            total_count = len(workbooks_data)
            unsaved_count = len([wb for wb in workbooks_data if not wb.get("saved", True)])

            if total_count == 1:
                message = "1개의 열린 워크북을 찾았습니다"
            else:
                message = f"{total_count}개의 열린 워크북을 찾았습니다"

            if has_unsaved:
                message += f" (저장되지 않은 워크북: {unsaved_count}개)"

            # 데이터 구성
            data_content = {
                "workbooks": workbooks_data,
                "total_count": total_count,
                "unsaved_count": unsaved_count,
                "has_unsaved": has_unsaved,
            }

            # 성공 응답 생성
            response = create_success_response(
                data=data_content, command="workbook-list", message=message, execution_time_ms=timer.execution_time_ms
            )

            # 출력 형식에 따른 결과 반환
            if output_format == "json":
                typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
            else:  # text 형식
                typer.echo(f"📊 {message}")
                typer.echo()

                if total_count == 0:
                    typer.echo("📋 열려있는 워크북이 없습니다.")
                    typer.echo("💡 Excel에서 워크북을 열거나 'oa excel workbook-open' 명령어를 사용하세요.")
                else:
                    for i, wb in enumerate(workbooks_data, 1):
                        status_icon = "💾" if wb.get("saved", True) else "⚠️"
                        typer.echo(f"{status_icon} {i}. {wb['name']}")

                        # 상세 정보 항상 표시
                        if "full_name" in wb:
                            typer.echo(f"   📁 경로: {wb['full_name']}")
                            typer.echo(f"   📄 시트 수: {wb['sheet_count']}")
                            typer.echo(f"   📑 활성 시트: {wb['active_sheet']}")

                            if "file_size_bytes" in wb:
                                size_mb = wb["file_size_bytes"] / (1024 * 1024)
                                typer.echo(f"   💽 파일 크기: {size_mb:.1f} MB")
                                typer.echo(f"   🕐 수정 시간: {wb['last_modified']}")

                        if not wb.get("saved", True):
                            typer.echo(f"   ⚠️  저장되지 않은 변경사항이 있습니다!")

                        if "error" in wb:
                            typer.echo(f"   ❌ {wb['error']}")

                        typer.echo()

    except Exception as e:
        error_response = create_error_response(e, "workbook-list")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ 오류 발생: {str(e)}", err=True)
            typer.echo("💡 Excel이 실행되고 있는지 확인하세요.", err=True)
        raise typer.Exit(1)


if __name__ == "__main__":
    typer.run(workbook_list)
