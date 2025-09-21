"""
피벗테이블 새로고침 명령어
데이터 소스 변경 사항을 피벗테이블에 반영
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


def pivot_refresh(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="피벗테이블이 있는 Excel 파일의 절대 경로"),
    use_active: bool = typer.Option(False, "--use-active", help="현재 활성 워크북 사용"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help='열린 워크북 이름으로 접근 (예: "Sales.xlsx")'),
    pivot_name: Optional[str] = typer.Option(
        None, "--pivot-name", help="새로고침할 피벗테이블 이름 (지정하지 않으면 전체 새로고침)"
    ),
    sheet: Optional[str] = typer.Option(None, "--sheet", help="피벗테이블이 있는 시트 이름 (지정하지 않으면 전체 워크북)"),
    refresh_all: bool = typer.Option(False, "--refresh-all", help="워크북의 모든 피벗테이블 새로고침 (기본값: False)"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택 (json/text)"),
    visible: bool = typer.Option(False, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부 (기본값: False)"),
    save: bool = typer.Option(True, "--save", help="새로고침 후 파일 저장 여부 (기본값: True)"),
):
    """
    피벗테이블의 데이터를 새로고침합니다.

    소스 데이터가 변경된 후 피벗테이블에 반영하기 위해 사용합니다.
    특정 피벗테이블 또는 전체 피벗테이블을 새로고침할 수 있습니다.

    워크북 접근 방법:
    - --file-path: 파일 경로로 워크북 열기
    - --use-active: 현재 활성 워크북 사용
    - --workbook-name: 열린 워크북 이름으로 접근

    예제:
        oa excel pivot-refresh --use-active --pivot-name "PivotTable1"
        oa excel pivot-refresh --file-path "sales.xlsx" --refresh-all
        oa excel pivot-refresh --workbook-name "Report.xlsx" --sheet "Dashboard"
    """
    book = None

    try:
        # 워크북 연결
        book = get_or_open_workbook(file_path=file_path, workbook_name=workbook_name, use_active=use_active, visible=visible)

        refresh_results = {
            "refreshed_pivots": [],
            "failed_pivots": [],
            "total_processed": 0,
            "success_count": 0,
            "error_count": 0,
        }

        # 플랫폼별 처리
        if platform.system() == "Windows":
            # Windows: COM API 사용

            if refresh_all:
                # 전체 워크북의 모든 피벗테이블 새로고침
                for ws in book.sheets:
                    try:
                        for pivot_table in ws.api.PivotTables():
                            pivot_info = {"name": pivot_table.Name, "sheet": ws.name, "status": "success"}
                            try:
                                pivot_table.RefreshTable()
                                refresh_results["refreshed_pivots"].append(pivot_info)
                                refresh_results["success_count"] += 1
                            except Exception as e:
                                pivot_info["status"] = "failed"
                                pivot_info["error"] = str(e)
                                refresh_results["failed_pivots"].append(pivot_info)
                                refresh_results["error_count"] += 1

                            refresh_results["total_processed"] += 1
                    except:
                        # 시트에 피벗테이블이 없거나 접근 불가
                        continue

            elif pivot_name:
                # 특정 피벗테이블 새로고침
                target_sheet = None
                pivot_table = None

                # 특정 시트가 지정된 경우
                if sheet:
                    target_sheet = get_sheet(book, sheet)
                    try:
                        pivot_table = target_sheet.api.PivotTables(pivot_name)
                    except:
                        raise ValueError(f"시트 '{sheet}'에서 피벗테이블 '{pivot_name}'을 찾을 수 없습니다")
                else:
                    # 전체 워크북에서 피벗테이블 검색
                    for ws in book.sheets:
                        try:
                            pivot_table = ws.api.PivotTables(pivot_name)
                            target_sheet = ws
                            break
                        except:
                            continue

                    if not pivot_table:
                        raise ValueError(f"피벗테이블 '{pivot_name}'을 찾을 수 없습니다")

                # 피벗테이블 새로고침
                pivot_info = {"name": pivot_name, "sheet": target_sheet.name, "status": "success"}

                try:
                    # 새로고침 전 정보 수집
                    refresh_date_before = None
                    try:
                        refresh_date_before = str(pivot_table.RefreshDate) if hasattr(pivot_table, "RefreshDate") else None
                    except:
                        pass

                    # 새로고침 실행
                    pivot_table.RefreshTable()

                    # 새로고침 후 정보 수집
                    refresh_date_after = None
                    try:
                        refresh_date_after = str(pivot_table.RefreshDate) if hasattr(pivot_table, "RefreshDate") else None
                    except:
                        pass

                    pivot_info.update({"refresh_date_before": refresh_date_before, "refresh_date_after": refresh_date_after})

                    refresh_results["refreshed_pivots"].append(pivot_info)
                    refresh_results["success_count"] = 1
                    refresh_results["total_processed"] = 1

                except Exception as e:
                    pivot_info["status"] = "failed"
                    pivot_info["error"] = str(e)
                    refresh_results["failed_pivots"].append(pivot_info)
                    refresh_results["error_count"] = 1
                    refresh_results["total_processed"] = 1

            elif sheet:
                # 특정 시트의 모든 피벗테이블 새로고침
                target_sheet = get_sheet(book, sheet)
                try:
                    for pivot_table in target_sheet.api.PivotTables():
                        pivot_info = {"name": pivot_table.Name, "sheet": target_sheet.name, "status": "success"}
                        try:
                            pivot_table.RefreshTable()
                            refresh_results["refreshed_pivots"].append(pivot_info)
                            refresh_results["success_count"] += 1
                        except Exception as e:
                            pivot_info["status"] = "failed"
                            pivot_info["error"] = str(e)
                            refresh_results["failed_pivots"].append(pivot_info)
                            refresh_results["error_count"] += 1

                        refresh_results["total_processed"] += 1
                except:
                    raise ValueError(f"시트 '{sheet}'에서 피벗테이블을 찾을 수 없습니다")

            else:
                raise ValueError("새로고침할 대상을 지정해주세요: --pivot-name, --sheet, 또는 --refresh-all")

        else:
            # macOS: 제한적 지원
            raise RuntimeError(
                "피벗테이블 새로고침은 Windows에서만 완전히 지원됩니다. macOS에서는 Excel의 수동 새로고침을 사용해주세요."
            )

        # 피벗캐시 새로고침도 시도 (선택적)
        if platform.system() == "Windows" and refresh_results["success_count"] > 0:
            try:
                # 워크북의 모든 피벗캐시 새로고침
                for cache_index in range(1, book.api.PivotCaches().Count + 1):
                    try:
                        book.api.PivotCaches(cache_index).Refresh()
                    except:
                        pass
                refresh_results["pivot_cache_refreshed"] = True
            except Exception as e:
                refresh_results["pivot_cache_error"] = str(e)

        # 파일 저장
        save_success = False
        save_error = None
        if save and refresh_results["success_count"] > 0:
            try:
                book.save()
                save_success = True
            except Exception as e:
                save_error = str(e)

        # 응답 데이터 구성
        data_content = {
            "refresh_results": refresh_results,
            "platform": platform.system(),
            "file_info": {
                "path": (
                    str(Path(normalize_path(file_path)).resolve())
                    if file_path
                    else (normalize_path(book.fullname) if hasattr(book, "fullname") else None)
                ),
                "name": Path(normalize_path(file_path)).name if file_path else normalize_path(book.name),
                "saved": save_success,
            },
        }

        if save_error:
            data_content["save_error"] = save_error

        # 성공 메시지 구성
        if refresh_results["success_count"] > 0:
            message = f"{refresh_results['success_count']}개 피벗테이블이 성공적으로 새로고침되었습니다"
            if refresh_results["error_count"] > 0:
                message += f" ({refresh_results['error_count']}개 실패)"
        else:
            message = "새로고침된 피벗테이블이 없습니다"

        response = create_success_response(data=data_content, command="pivot-refresh", message=message)

        # 출력 형식에 따른 결과 반환
        if output_format == "json":
            typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
        else:  # text 형식
            typer.echo(f"✅ 피벗테이블 새로고침 완료")
            typer.echo(f"📄 파일: {data_content['file_info']['name']}")
            typer.echo(f"📊 처리된 피벗테이블: {refresh_results['total_processed']}개")
            typer.echo(f"✅ 성공: {refresh_results['success_count']}개")

            if refresh_results["error_count"] > 0:
                typer.echo(f"❌ 실패: {refresh_results['error_count']}개")

            typer.echo()

            # 성공한 피벗테이블들 표시
            if refresh_results["refreshed_pivots"]:
                typer.echo("✅ 새로고침 성공:")
                for pivot in refresh_results["refreshed_pivots"]:
                    typer.echo(f"   📋 {pivot['name']} ({pivot['sheet']})")
                    if pivot.get("refresh_date_after"):
                        typer.echo(f"      🕒 새로고침 시간: {pivot['refresh_date_after']}")

            # 실패한 피벗테이블들 표시
            if refresh_results["failed_pivots"]:
                typer.echo("\n❌ 새로고침 실패:")
                for pivot in refresh_results["failed_pivots"]:
                    typer.echo(f"   📋 {pivot['name']} ({pivot['sheet']})")
                    typer.echo(f"      ❌ 오류: {pivot['error']}")

            # 피벗캐시 정보
            if refresh_results.get("pivot_cache_refreshed"):
                typer.echo("\n🔄 피벗캐시도 새로고침되었습니다")
            elif refresh_results.get("pivot_cache_error"):
                typer.echo(f"\n⚠️ 피벗캐시 새로고침 실패: {refresh_results['pivot_cache_error']}")

            if save_success:
                typer.echo("\n💾 파일이 저장되었습니다")
            elif save and refresh_results["success_count"] > 0:
                typer.echo(f"\n⚠️ 저장 실패: {save_error}")
            elif refresh_results["success_count"] > 0:
                typer.echo("\n📝 파일이 저장되지 않았습니다 (--save=False)")

    except ValueError as e:
        error_response = create_error_response(e, "pivot-refresh")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)

    except RuntimeError as e:
        error_response = create_error_response(e, "pivot-refresh")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
            if "Windows" in str(e):
                typer.echo(
                    "💡 피벗테이블 새로고침은 Windows에서만 완전히 지원됩니다. macOS에서는 Excel의 수동 기능을 사용해주세요.",
                    err=True,
                )
            else:
                typer.echo(
                    "💡 Excel이 설치되어 있는지 확인하고, 파일이 다른 프로그램에서 사용 중이지 않은지 확인하세요.", err=True
                )
        raise typer.Exit(1)

    except Exception as e:
        error_response = create_error_response(e, "pivot-refresh")
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
