"""
Excel 워크북 목록 조회 명령어 (Typer 버전)
현재 열려있는 모든 워크북들의 목록과 기본 정보 제공
"""

import datetime
import json
import sys
from pathlib import Path
from typing import Optional

import typer
import xlwings as xw

from pyhub_office_automation.version import get_version

from .utils import ExecutionTimer, create_error_response, create_success_response, normalize_path


def workbook_list(
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택"),
    detailed: bool = typer.Option(False, "--detailed", help="상세 정보 포함 (파일 경로, 시트 수, 저장 상태 등)"),
):
    """
    현재 열려있는 모든 Excel 워크북 목록을 조회합니다.

    기본적으로 워크북 이름만 반환하며, --detailed 옵션으로 상세 정보를 포함할 수 있습니다.

    예제:
        oa excel workbook-list
        oa excel workbook-list --detailed --format text
    """
    try:
        # 실행 시간 측정 시작
        with ExecutionTimer() as timer:
            # 현재 열린 워크북들 확인
            if len(xw.books) == 0:
                # 열린 워크북이 없는 경우
                workbooks_data = []
                has_unsaved = False
                message = "현재 열려있는 워크북이 없습니다"
            else:
                workbooks_data = []
                has_unsaved = False

                for book in xw.books:
                    try:
                        # 안전하게 saved 상태 확인
                        try:
                            saved_status = book.saved
                        except:
                            saved_status = True  # 기본값으로 저장됨으로 가정

                        workbook_info = {"name": normalize_path(book.name), "saved": saved_status}

                        # 저장되지 않은 워크북 체크
                        if not saved_status:
                            has_unsaved = True

                        if detailed:
                            # 상세 정보 추가
                            workbook_info.update(
                                {
                                    "full_name": normalize_path(book.fullname),
                                    "sheet_count": len(book.sheets),
                                    "active_sheet": book.sheets.active.name if book.sheets else None,
                                }
                            )

                            # 파일 정보 추가 (파일이 실제로 존재하는 경우)
                            try:
                                file_path = Path(book.fullname)
                                if file_path.exists():
                                    file_stat = file_path.stat()
                                    workbook_info.update(
                                        {
                                            "file_size_bytes": file_stat.st_size,
                                            "last_modified": datetime.datetime.fromtimestamp(file_stat.st_mtime).isoformat(),
                                        }
                                    )
                            except (OSError, AttributeError):
                                # 새 워크북이거나 파일 접근 불가능한 경우
                                pass

                        workbooks_data.append(workbook_info)

                    except Exception as e:
                        # 개별 워크북 정보 수집 실패 시 기본 정보만 포함
                        workbooks_data.append(
                            {
                                "name": getattr(book, "name", "Unknown"),
                                "saved": getattr(book, "saved", False),
                                "error": f"정보 수집 실패: {str(e)}",
                            }
                        )

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

                        if detailed and "full_name" in wb:
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
