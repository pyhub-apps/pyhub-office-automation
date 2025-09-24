"""
Excel 워크북 전체 메타데이터 자동 생성 명령어 (Issue #59)
워크북의 모든 Excel Table에 대한 메타데이터를 일괄 생성 및 저장
"""

import json
import platform
from pathlib import Path
from typing import Optional

import typer
import xlwings as xw

from pyhub_office_automation.version import get_version

from .metadata_utils import auto_generate_table_metadata, ensure_metadata_sheet, get_metadata_record, write_metadata_record
from .utils import ExecutionTimer, create_error_response, create_success_response, get_or_open_workbook, normalize_path


def metadata_generate(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="Excel 파일의 절대 경로"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help="열린 워크북 이름으로 접근"),
    all_tables: bool = typer.Option(True, "--all-tables/--no-all-tables", help="모든 Table 처리 여부"),
    specific_sheet: Optional[str] = typer.Option(None, "--sheet", help="특정 시트의 Table만 처리"),
    force_overwrite: bool = typer.Option(False, "--force-overwrite", help="기존 메타데이터가 있어도 강제 덮어쓰기"),
    skip_existing: bool = typer.Option(
        True, "--skip-existing/--no-skip-existing", help="기존 메타데이터가 있는 Table 건너뛰기"
    ),
    dry_run: bool = typer.Option(False, "--dry-run", help="실제 저장 없이 분석만 수행 (미리보기)"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택"),
    visible: bool = typer.Option(False, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부"),
):
    """
    워크북의 모든 Excel Table에 대한 메타데이터를 자동으로 생성합니다.

    각 Table을 분석하여 데이터 타입, 구조, 설명 등의 메타데이터를
    일괄적으로 생성하고 Metadata 시트에 저장합니다.

    \b
    워크북 접근 방법:
      • 옵션 없음: 활성 워크북 자동 사용 (기본값)
      • --file-path: 파일 경로로 워크북 열기
      • --workbook-name: 열린 워크북 이름으로 접근

    \b
    처리 범위 옵션:
      • --all-tables: 모든 시트의 모든 Table 처리 (기본값)
      • --sheet: 특정 시트의 Table만 처리
      • --force-overwrite: 기존 메타데이터 덮어쓰기
      • --skip-existing: 기존 메타데이터가 있는 경우 건너뛰기 (기본값)

    \b
    안전 옵션:
      • --dry-run: 실제 저장 없이 분석 결과만 확인
      • --no-all-tables: 명시적으로 지정된 Table만 처리

    \b
    사용 예제:
      # 전체 워크북 메타데이터 생성
      oa excel metadata-generate

      # 특정 파일의 메타데이터 생성
      oa excel metadata-generate --file-path "sales.xlsx"

      # 특정 시트만 처리
      oa excel metadata-generate --sheet "DataSheet"

      # 기존 데이터 덮어쓰기
      oa excel metadata-generate --force-overwrite

      # 미리보기 (실제 저장 안함)
      oa excel metadata-generate --dry-run
    """
    book = None
    try:
        with ExecutionTimer() as timer:
            # 플랫폼 확인
            if platform.system() != "Windows":
                typer.echo("⚠️ Excel Table 메타데이터 생성은 Windows에서 완전히 지원됩니다.")

            # 워크북 연결
            book = get_or_open_workbook(file_path=file_path, workbook_name=workbook_name, visible=visible)

            # Metadata 시트 확보 (dry_run이 아닌 경우만)
            if not dry_run:
                metadata_sheet = ensure_metadata_sheet(book)

            # 처리할 시트 결정
            if specific_sheet:
                try:
                    target_sheets = [book.sheets[specific_sheet]]
                except:
                    raise ValueError(f"시트 '{specific_sheet}'을 찾을 수 없습니다")
            else:
                target_sheets = list(book.sheets)

            # 모든 Table 수집
            all_found_tables = []
            processing_summary = {
                "total_sheets_scanned": len(target_sheets),
                "total_tables_found": 0,
                "tables_processed": 0,
                "tables_skipped": 0,
                "tables_failed": 0,
                "tables_created": 0,
                "tables_updated": 0,
            }

            processing_details = []

            for sheet in target_sheets:
                sheet_tables = []

                try:
                    if platform.system() == "Windows":
                        # Windows에서 COM API로 Table 조회
                        for table in sheet.api.ListObjects():
                            table_info = {
                                "name": table.Name,
                                "sheet": sheet.name,
                                "range": table.Range.Address.replace("$", ""),
                                "row_count": table.Range.Rows.Count - 1,  # 헤더 제외
                                "column_count": table.Range.Columns.Count,
                            }
                            sheet_tables.append(table_info)
                            all_found_tables.append(table_info)
                    else:
                        # macOS에서는 제한적인 지원
                        for table in sheet.tables:
                            table_info = {
                                "name": table.name,
                                "sheet": sheet.name,
                                "range": table.range.address.replace("$", ""),
                                "row_count": table.range.rows.count - 1,
                                "column_count": table.range.columns.count,
                            }
                            sheet_tables.append(table_info)
                            all_found_tables.append(table_info)

                except Exception as e:
                    # 시트 접근 실패 시 경고하고 계속 진행
                    typer.echo(f"⚠️ 시트 '{sheet.name}' 접근 실패: {str(e)}", err=True)
                    continue

            processing_summary["total_tables_found"] = len(all_found_tables)

            if not all_found_tables:
                message = f"처리할 Excel Table을 찾을 수 없습니다"
                if specific_sheet:
                    message += f" (시트: {specific_sheet})"

                data_content = {
                    "summary": processing_summary,
                    "processing_details": [],
                    "workbook": {
                        "name": normalize_path(book.name),
                        "full_name": normalize_path(book.fullname),
                        "saved": getattr(book, "saved", True),
                    },
                    "options": {
                        "all_tables": all_tables,
                        "specific_sheet": specific_sheet,
                        "force_overwrite": force_overwrite,
                        "skip_existing": skip_existing,
                        "dry_run": dry_run,
                    },
                }

                response = create_success_response(
                    data=data_content,
                    command="metadata-generate",
                    message=message,
                    execution_time_ms=timer.execution_time_ms,
                    book=book,
                )

                if output_format == "json":
                    typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
                else:
                    typer.echo(f"ℹ️ {message}")

                return

            # 각 Table 처리
            for table_info in all_found_tables:
                table_name = table_info["name"]
                sheet_name = table_info["sheet"]

                process_detail = {
                    "table_name": table_name,
                    "sheet_name": sheet_name,
                    "action": "none",
                    "success": False,
                    "message": "",
                    "metadata": None,
                }

                try:
                    # 기존 메타데이터 확인
                    existing_metadata = get_metadata_record(book, table_name)

                    # 처리 여부 결정
                    should_process = True
                    if existing_metadata:
                        if skip_existing and not force_overwrite:
                            process_detail["action"] = "skipped"
                            process_detail["message"] = "기존 메타데이터 존재 (건너뜀)"
                            process_detail["success"] = True
                            processing_summary["tables_skipped"] += 1
                            should_process = False
                        elif force_overwrite:
                            process_detail["action"] = "update"
                        else:
                            process_detail["action"] = "skipped"
                            process_detail["message"] = "기존 메타데이터 존재 (덮어쓰기 안함)"
                            process_detail["success"] = True
                            processing_summary["tables_skipped"] += 1
                            should_process = False
                    else:
                        process_detail["action"] = "create"

                    if should_process:
                        # Table 메타데이터 자동 생성
                        analysis_result = auto_generate_table_metadata(book, table_name, sheet_name)

                        if not analysis_result.get("success"):
                            process_detail["message"] = analysis_result.get("notes", "분석 실패")
                            processing_summary["tables_failed"] += 1
                        else:
                            process_detail["metadata"] = analysis_result

                            # 실제 저장 (dry_run이 아닌 경우만)
                            if not dry_run:
                                save_success = write_metadata_record(
                                    workbook=book,
                                    table_name=table_name,
                                    sheet_name=sheet_name,
                                    description=analysis_result["description"],
                                    data_type=analysis_result["data_type"],
                                    column_info=analysis_result["column_info"],
                                    row_count=analysis_result["row_count"],
                                    tags=analysis_result["tags"],
                                    notes=analysis_result["notes"],
                                )

                                if save_success:
                                    process_detail["success"] = True
                                    process_detail["message"] = "메타데이터 생성 및 저장 성공"
                                    if process_detail["action"] == "create":
                                        processing_summary["tables_created"] += 1
                                    else:
                                        processing_summary["tables_updated"] += 1
                                    processing_summary["tables_processed"] += 1
                                else:
                                    process_detail["message"] = "분석 성공, 저장 실패"
                                    processing_summary["tables_failed"] += 1
                            else:
                                process_detail["success"] = True
                                process_detail["message"] = "분석 완료 (dry-run 모드)"
                                processing_summary["tables_processed"] += 1

                except Exception as e:
                    process_detail["message"] = f"처리 실패: {str(e)}"
                    processing_summary["tables_failed"] += 1

                processing_details.append(process_detail)

            # 워크북 정보
            workbook_info = {
                "name": normalize_path(book.name),
                "full_name": normalize_path(book.fullname),
                "saved": getattr(book, "saved", True),
                "total_sheets": len(book.sheets),
            }

            # 결과 데이터 구성
            data_content = {
                "summary": processing_summary,
                "processing_details": processing_details,
                "workbook": workbook_info,
                "options": {
                    "all_tables": all_tables,
                    "specific_sheet": specific_sheet,
                    "force_overwrite": force_overwrite,
                    "skip_existing": skip_existing,
                    "dry_run": dry_run,
                },
            }

            # 성공 메시지 생성
            processed = processing_summary["tables_processed"]
            skipped = processing_summary["tables_skipped"]
            failed = processing_summary["tables_failed"]
            total = processing_summary["total_tables_found"]

            status_parts = []
            if processed > 0:
                status_parts.append(f"{processed}개 처리")
            if skipped > 0:
                status_parts.append(f"{skipped}개 건너뜀")
            if failed > 0:
                status_parts.append(f"{failed}개 실패")

            status_str = ", ".join(status_parts) if status_parts else "처리 없음"

            dry_run_suffix = " (미리보기 모드)" if dry_run else ""
            message = f"메타데이터 생성 완료: 총 {total}개 테이블 중 {status_str}{dry_run_suffix}"

            # 성공 응답 생성
            response = create_success_response(
                data=data_content,
                command="metadata-generate",
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

                # 처리 요약
                typer.echo("📊 처리 요약:")
                typer.echo(f"  📁 워크북: {workbook_info['name']}")
                typer.echo(f"  📄 스캔한 시트: {processing_summary['total_sheets_scanned']}개")
                typer.echo(f"  🏷️ 발견한 테이블: {processing_summary['total_tables_found']}개")
                typer.echo(f"  ✅ 처리된 테이블: {processing_summary['tables_processed']}개")
                if processing_summary["tables_created"] > 0:
                    typer.echo(f"    └ 새로 생성: {processing_summary['tables_created']}개")
                if processing_summary["tables_updated"] > 0:
                    typer.echo(f"    └ 업데이트: {processing_summary['tables_updated']}개")
                if processing_summary["tables_skipped"] > 0:
                    typer.echo(f"  ⏭️ 건너뛴 테이블: {processing_summary['tables_skipped']}개")
                if processing_summary["tables_failed"] > 0:
                    typer.echo(f"  ❌ 실패한 테이블: {processing_summary['tables_failed']}개")

                # 상세 결과 (실패한 것만 표시)
                failed_details = [d for d in processing_details if not d["success"]]
                if failed_details:
                    typer.echo()
                    typer.echo("❌ 실패한 테이블:")
                    for detail in failed_details:
                        typer.echo(f"  • {detail['table_name']} ({detail['sheet_name']}): {detail['message']}")

                if dry_run:
                    typer.echo()
                    typer.echo("💡 --dry-run 모드입니다. 실제로 저장하려면 이 옵션을 제거하고 다시 실행하세요.")

    except FileNotFoundError as e:
        error_response = create_error_response(e, "metadata-generate")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ 파일을 찾을 수 없습니다", err=True)
        raise typer.Exit(1)

    except ValueError as e:
        error_response = create_error_response(e, "metadata-generate")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)

    except Exception as e:
        error_response = create_error_response(e, "metadata-generate")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ 예기치 않은 오류: {str(e)}", err=True)
            typer.echo("💡 Excel이 설치되어 있는지 확인하고, 워크북에 Excel Table이 있는지 확인하세요.", err=True)
        raise typer.Exit(1)

    finally:
        # 워크북 정리 - 활성 워크북이나 이름으로 접근한 경우 앱 종료하지 않음
        if book and not visible and file_path:
            try:
                book.app.quit()
            except:
                pass


if __name__ == "__main__":
    typer.run(metadata_generate)
