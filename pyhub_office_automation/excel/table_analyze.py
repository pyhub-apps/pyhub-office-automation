"""
Excel Table 분석 및 메타데이터 자동 생성 명령어 (Issue #59)
특정 Table의 메타데이터를 자동 분석하고 Metadata 시트에 저장
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
from .metadata_utils import (
    auto_generate_table_metadata,
    write_metadata_record,
    get_metadata_record,
)


def table_analyze(
    table_name: str = typer.Option(..., "--table-name", help="분석할 Excel Table 이름"),
    file_path: Optional[str] = typer.Option(None, "--file-path", help="Excel 파일의 절대 경로"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help="열린 워크북 이름으로 접근"),
    sheet: Optional[str] = typer.Option(None, "--sheet", help="시트 이름 (미지정시 Table 검색으로 자동 찾기)"),
    update_metadata: bool = typer.Option(True, "--update-metadata/--no-update-metadata", help="Metadata 시트에 결과 저장 여부"),
    force_overwrite: bool = typer.Option(False, "--force-overwrite", help="기존 메타데이터가 있어도 강제 덮어쓰기"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택"),
    visible: bool = typer.Option(False, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부"),
):
    """
    Excel Table을 분석하고 메타데이터를 자동으로 생성합니다.

    Table의 구조, 데이터 타입, 행 수 등을 자동으로 분석하여
    Metadata 시트에 저장하거나 JSON으로 출력합니다.

    \b
    워크북 접근 방법:
      • 옵션 없음: 활성 워크북 자동 사용 (기본값)
      • --file-path: 파일 경로로 워크북 열기
      • --workbook-name: 열린 워크북 이름으로 접근

    \b
    분석 내용:
      • 테이블 기본 정보 (행/열 수, 범위)
      • 컬럼 구조 및 헤더 분석
      • 데이터 타입 추론 (sales, customer, product 등)
      • 자동 태그 생성 (large-dataset, auto-generated 등)
      • 비즈니스 설명 자동 생성

    \b
    사용 예제:
      # 활성 워크북의 특정 Table 분석
      oa excel table-analyze --table-name "SalesData"

      # 특정 파일의 Table 분석 후 메타데이터 저장
      oa excel table-analyze --table-name "ProductList" --file-path "inventory.xlsx"

      # 기존 메타데이터 강제 덮어쓰기
      oa excel table-analyze --table-name "CustomerData" --force-overwrite

      # 분석만 하고 저장하지 않음
      oa excel table-analyze --table-name "TempData" --no-update-metadata
    """
    book = None
    try:
        with ExecutionTimer() as timer:
            # 플랫폼 확인
            if platform.system() != "Windows":
                typer.echo("⚠️ Excel Table 분석은 Windows에서 완전히 지원됩니다. macOS에서는 제한된 기능만 사용 가능합니다.")

            # 워크북 연결
            book = get_or_open_workbook(file_path=file_path, workbook_name=workbook_name, visible=visible)

            # Table이 있는 시트 찾기
            target_sheet = None
            target_sheet_name = None

            if sheet:
                # 지정된 시트에서 Table 찾기
                try:
                    target_sheet = get_sheet(book, sheet)
                    target_sheet_name = sheet

                    # 해당 시트에 Table이 있는지 확인
                    table_found = False
                    if platform.system() == "Windows":
                        for table in target_sheet.api.ListObjects():
                            if table.Name == table_name:
                                table_found = True
                                break
                    else:
                        for table in target_sheet.tables:
                            if table.name == table_name:
                                table_found = True
                                break

                    if not table_found:
                        raise ValueError(f"시트 '{sheet}'에서 테이블 '{table_name}'을 찾을 수 없습니다")

                except Exception as e:
                    raise ValueError(f"시트 '{sheet}' 접근 실패: {str(e)}")
            else:
                # 모든 시트에서 Table 검색
                for ws in book.sheets:
                    try:
                        if platform.system() == "Windows":
                            for table in ws.api.ListObjects():
                                if table.Name == table_name:
                                    target_sheet = ws
                                    target_sheet_name = ws.name
                                    break
                        else:
                            for table in ws.tables:
                                if table.name == table_name:
                                    target_sheet = ws
                                    target_sheet_name = ws.name
                                    break
                    except:
                        continue

                    if target_sheet:
                        break

                if not target_sheet:
                    raise ValueError(f"워크북에서 테이블 '{table_name}'을 찾을 수 없습니다. --sheet 옵션으로 시트를 지정해보세요.")

            # 기존 메타데이터 확인
            existing_metadata = get_metadata_record(book, table_name)
            if existing_metadata and not force_overwrite:
                if update_metadata:
                    typer.echo(f"⚠️ 테이블 '{table_name}'의 메타데이터가 이미 존재합니다. --force-overwrite 옵션을 사용하여 덮어쓰기하거나 --no-update-metadata로 분석만 수행하세요.")

            # Table 메타데이터 자동 생성
            analysis_result = auto_generate_table_metadata(book, table_name, target_sheet_name)

            if not analysis_result.get("success"):
                raise ValueError(analysis_result.get("notes", f"테이블 '{table_name}' 분석 실패"))

            # Metadata 시트에 저장
            saved_to_metadata = False
            if update_metadata and (not existing_metadata or force_overwrite):
                save_success = write_metadata_record(
                    workbook=book,
                    table_name=table_name,
                    sheet_name=target_sheet_name,
                    description=analysis_result["description"],
                    data_type=analysis_result["data_type"],
                    column_info=analysis_result["column_info"],
                    row_count=analysis_result["row_count"],
                    tags=analysis_result["tags"],
                    notes=analysis_result["notes"]
                )
                saved_to_metadata = save_success

            # 워크북 정보
            workbook_info = {
                "name": normalize_path(book.name),
                "full_name": normalize_path(book.fullname),
                "saved": getattr(book, "saved", True),
            }

            # 결과 데이터 구성
            data_content = {
                "table_name": table_name,
                "sheet_name": target_sheet_name,
                "analysis_result": analysis_result,
                "metadata_action": {
                    "saved_to_metadata": saved_to_metadata,
                    "overwritten": force_overwrite and existing_metadata is not None,
                    "skipped_reason": None if saved_to_metadata else ("existing_metadata" if existing_metadata else "update_disabled")
                },
                "workbook": workbook_info,
                "options": {
                    "update_metadata": update_metadata,
                    "force_overwrite": force_overwrite,
                }
            }

            # 성공 메시지 생성
            if saved_to_metadata:
                action_msg = "분석 완료 및 메타데이터 저장됨"
                if force_overwrite and existing_metadata:
                    action_msg += " (기존 데이터 덮어씀)"
            elif existing_metadata and update_metadata:
                action_msg = "분석 완료 (기존 메타데이터 유지, --force-overwrite로 덮어쓰기 가능)"
            elif not update_metadata:
                action_msg = "분석 완료 (메타데이터 저장 안함)"
            else:
                action_msg = "분석 완료 (메타데이터 저장 실패)"

            message = f"테이블 '{table_name}' ({target_sheet_name} 시트) {action_msg}"

            # 성공 응답 생성
            response = create_success_response(
                data=data_content,
                command="table-analyze",
                message=message,
                execution_time_ms=timer.execution_time_ms,
                book=book,
            )

            # 출력 형식에 따른 결과 반환
            if output_format == "json":
                typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
            else:  # text 형식
                typer.echo(f"✅ {message}")
                typer.echo()

                # 분석 결과 요약
                analysis = analysis_result
                typer.echo("📊 분석 결과:")
                typer.echo(f"  🏷️ 테이블: {table_name}")
                typer.echo(f"  📄 시트: {target_sheet_name}")
                typer.echo(f"  📝 설명: {analysis['description']}")
                typer.echo(f"  🏷️ 데이터 타입: {analysis['data_type']}")
                typer.echo(f"  📊 크기: {analysis['row_count']}행")
                typer.echo(f"  📋 컬럼: {analysis['column_info']}")
                typer.echo(f"  🏷️ 태그: {analysis['tags']}")

                if saved_to_metadata:
                    typer.echo()
                    typer.echo("💾 Metadata 시트에 저장되었습니다.")
                elif existing_metadata:
                    typer.echo()
                    typer.echo("⚠️ 기존 메타데이터가 있습니다. --force-overwrite로 덮어쓸 수 있습니다.")

    except FileNotFoundError as e:
        error_response = create_error_response(e, "table-analyze")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ 파일을 찾을 수 없습니다", err=True)
        raise typer.Exit(1)

    except ValueError as e:
        error_response = create_error_response(e, "table-analyze")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)

    except Exception as e:
        error_response = create_error_response(e, "table-analyze")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ 예기치 않은 오류: {str(e)}", err=True)
            typer.echo("💡 Excel이 설치되어 있는지 확인하고, 테이블 이름을 정확히 입력했는지 확인하세요.", err=True)
        raise typer.Exit(1)

    finally:
        # 워크북 정리 - 활성 워크북이나 이름으로 접근한 경우 앱 종료하지 않음
        if book and not visible and file_path:
            try:
                book.app.quit()
            except:
                pass


if __name__ == "__main__":
    typer.run(table_analyze)