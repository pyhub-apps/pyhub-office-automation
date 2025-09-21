"""
피벗테이블 삭제 명령어
워크북에서 특정 피벗테이블을 삭제
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


def pivot_delete(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="피벗테이블이 있는 Excel 파일의 절대 경로"),
    use_active: bool = typer.Option(False, "--use-active", help="현재 활성 워크북 사용"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help='열린 워크북 이름으로 접근 (예: "Sales.xlsx")'),
    pivot_name: str = typer.Option(..., "--pivot-name", help="삭제할 피벗테이블 이름"),
    sheet: Optional[str] = typer.Option(None, "--sheet", help="피벗테이블이 있는 시트 이름 (지정하지 않으면 자동 검색)"),
    confirm: bool = typer.Option(False, "--confirm", help="삭제 확인 (기본값: False, True로 설정해야 실제 삭제)"),
    delete_cache: bool = typer.Option(False, "--delete-cache", help="연관된 피벗캐시도 삭제 (기본값: False)"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택 (json/text)"),
    visible: bool = typer.Option(False, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부 (기본값: False)"),
    save: bool = typer.Option(True, "--save", help="삭제 후 파일 저장 여부 (기본값: True)"),
):
    """
    지정된 피벗테이블을 삭제합니다.

    안전을 위해 --confirm=True 옵션을 지정해야 실제로 삭제됩니다.
    Windows 전용 기능입니다.

    워크북 접근 방법:
    - --file-path: 파일 경로로 워크북 열기
    - --use-active: 현재 활성 워크북 사용
    - --workbook-name: 열린 워크북 이름으로 접근

    예제:
        oa excel pivot-delete --use-active --pivot-name "PivotTable1" --confirm=True
        oa excel pivot-delete --file-path "sales.xlsx" --pivot-name "SalesPivot" --confirm=True --delete-cache=True
        oa excel pivot-delete --workbook-name "Report.xlsx" --pivot-name "Dashboard" --sheet "Data" --confirm=True
    """
    book = None

    try:
        # Windows 전용 기능 확인
        if platform.system() != "Windows":
            raise RuntimeError("피벗테이블 삭제는 Windows에서만 지원됩니다. macOS에서는 수동으로 피벗테이블을 삭제해주세요.")

        # 삭제 확인
        if not confirm:
            raise ValueError("안전을 위해 --confirm=True 옵션을 지정해야 피벗테이블이 삭제됩니다")

        # 워크북 연결
        book = get_or_open_workbook(file_path=file_path, workbook_name=workbook_name, use_active=use_active, visible=visible)

        # 피벗테이블 찾기
        target_sheet = None
        pivot_table = None
        pivot_info = None

        # 특정 시트가 지정된 경우
        if sheet:
            target_sheet = get_sheet(book, sheet)
            try:
                pivot_table = target_sheet.api.PivotTables(pivot_name)
                pivot_info = {
                    "name": pivot_table.Name,
                    "sheet": target_sheet.name,
                    "location": pivot_table.TableRange1.Address if hasattr(pivot_table, "TableRange1") else "Unknown",
                }
            except:
                raise ValueError(f"시트 '{sheet}'에서 피벗테이블 '{pivot_name}'을 찾을 수 없습니다")
        else:
            # 전체 워크북에서 피벗테이블 검색
            for ws in book.sheets:
                try:
                    pivot_table = ws.api.PivotTables(pivot_name)
                    target_sheet = ws
                    pivot_info = {
                        "name": pivot_table.Name,
                        "sheet": target_sheet.name,
                        "location": pivot_table.TableRange1.Address if hasattr(pivot_table, "TableRange1") else "Unknown",
                    }
                    break
                except:
                    continue

            if not pivot_table:
                raise ValueError(f"피벗테이블 '{pivot_name}'을 찾을 수 없습니다")

        # 삭제 전 정보 수집
        try:
            # 피벗테이블 상세 정보 수집
            pivot_info.update(
                {
                    "source_data": pivot_table.SourceData if hasattr(pivot_table, "SourceData") else "Unknown",
                    "cache_index": pivot_table.CacheIndex if hasattr(pivot_table, "CacheIndex") else None,
                    "refresh_date": str(pivot_table.RefreshDate) if hasattr(pivot_table, "RefreshDate") else None,
                }
            )

            # 관련 필드 정보 수집
            field_info = {"row_fields": [], "column_fields": [], "data_fields": [], "page_fields": []}

            try:
                field_info["row_fields"] = [field.Name for field in pivot_table.RowFields]
            except:
                pass

            try:
                field_info["column_fields"] = [field.Name for field in pivot_table.ColumnFields]
            except:
                pass

            try:
                field_info["data_fields"] = [field.Name for field in pivot_table.DataFields]
            except:
                pass

            try:
                field_info["page_fields"] = [field.Name for field in pivot_table.PageFields]
            except:
                pass

            pivot_info["fields"] = field_info

        except Exception as e:
            pivot_info["info_collection_error"] = f"정보 수집 중 오류: {str(e)}"

        # 피벗캐시 정보 수집 (삭제할 경우를 위해)
        cache_info = None
        if delete_cache and pivot_info.get("cache_index"):
            try:
                cache_index = pivot_info["cache_index"]
                pivot_cache = book.api.PivotCaches(cache_index)
                cache_info = {
                    "index": cache_index,
                    "source_data": pivot_cache.SourceData if hasattr(pivot_cache, "SourceData") else "Unknown",
                }
            except Exception as e:
                cache_info = {"error": f"캐시 정보 수집 실패: {str(e)}"}

        # 피벗테이블 삭제 실행
        delete_results = {"pivot_deleted": False, "cache_deleted": False, "errors": []}

        try:
            # 피벗테이블 삭제
            pivot_table.TableRange2.Delete() if hasattr(pivot_table, "TableRange2") else pivot_table.TableRange1.Delete()
            delete_results["pivot_deleted"] = True

        except Exception as e:
            delete_results["errors"].append(f"피벗테이블 삭제 실패: {str(e)}")

        # 피벗캐시 삭제 (선택적)
        if delete_cache and cache_info and not cache_info.get("error"):
            try:
                cache_index = cache_info["index"]

                # 해당 캐시를 사용하는 다른 피벗테이블이 있는지 확인
                cache_in_use = False
                for ws in book.sheets:
                    try:
                        for pt in ws.api.PivotTables():
                            if hasattr(pt, "CacheIndex") and pt.CacheIndex == cache_index:
                                cache_in_use = True
                                break
                    except:
                        continue
                    if cache_in_use:
                        break

                if not cache_in_use:
                    # 캐시를 사용하는 피벗테이블이 없으면 삭제
                    book.api.PivotCaches(cache_index).Delete()
                    delete_results["cache_deleted"] = True
                else:
                    delete_results["errors"].append("피벗캐시가 다른 피벗테이블에서 사용 중이므로 삭제하지 않았습니다")

            except Exception as e:
                delete_results["errors"].append(f"피벗캐시 삭제 실패: {str(e)}")

        # 삭제 성공 여부 확인
        if not delete_results["pivot_deleted"]:
            raise RuntimeError("피벗테이블 삭제에 실패했습니다")

        # 파일 저장
        save_success = False
        save_error = None
        if save:
            try:
                book.save()
                save_success = True
            except Exception as e:
                save_error = str(e)

        # 응답 데이터 구성
        data_content = {
            "deleted_pivot": pivot_info,
            "delete_results": delete_results,
            "cache_info": cache_info,
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
        message = f"피벗테이블 '{pivot_name}'이 성공적으로 삭제되었습니다"
        if delete_results.get("cache_deleted"):
            message += " (피벗캐시 포함)"

        response = create_success_response(data=data_content, command="pivot-delete", message=message)

        # 출력 형식에 따른 결과 반환
        if output_format == "json":
            typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
        else:  # text 형식
            typer.echo(f"✅ 피벗테이블 삭제 완료")
            typer.echo(f"📋 피벗테이블 이름: {pivot_name}")
            typer.echo(f"📄 파일: {data_content['file_info']['name']}")
            typer.echo(f"📍 시트: {target_sheet.name}")
            typer.echo(f"📍 위치: {pivot_info.get('location', 'Unknown')}")

            # 삭제된 필드 정보 표시
            if pivot_info.get("fields"):
                fields = pivot_info["fields"]
                field_summary = []
                if fields.get("row_fields"):
                    field_summary.append(f"행: {', '.join(fields['row_fields'])}")
                if fields.get("column_fields"):
                    field_summary.append(f"열: {', '.join(fields['column_fields'])}")
                if fields.get("data_fields"):
                    field_summary.append(f"값: {', '.join(fields['data_fields'])}")
                if fields.get("page_fields"):
                    field_summary.append(f"필터: {', '.join(fields['page_fields'])}")

                if field_summary:
                    typer.echo(f"📊 삭제된 필드: {' | '.join(field_summary)}")

            # 캐시 정보
            if delete_results.get("cache_deleted"):
                typer.echo("🗑️ 연관된 피벗캐시도 삭제되었습니다")
            elif cache_info and not cache_info.get("error"):
                typer.echo("💾 피벗캐시는 다른 피벗테이블에서 사용 중이므로 보존되었습니다")

            # 오류 표시
            if delete_results.get("errors"):
                typer.echo("\n⚠️ 삭제 과정에서 발생한 경고:")
                for error in delete_results["errors"]:
                    typer.echo(f"   {error}")

            if save_success:
                typer.echo("\n💾 파일이 저장되었습니다")
            elif save:
                typer.echo(f"\n⚠️ 저장 실패: {save_error}")
            else:
                typer.echo("\n📝 파일이 저장되지 않았습니다 (--save=False)")

            typer.echo("\n💡 피벗테이블 목록 확인은 'oa excel pivot-list' 명령어를 사용하세요")

    except ValueError as e:
        error_response = create_error_response(e, "pivot-delete")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
            if "confirm" in str(e).lower():
                typer.echo("💡 안전을 위해 --confirm=True 옵션을 반드시 지정해야 합니다", err=True)
        raise typer.Exit(1)

    except RuntimeError as e:
        error_response = create_error_response(e, "pivot-delete")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
            if "Windows" in str(e):
                typer.echo(
                    "💡 피벗테이블 삭제는 Windows에서만 지원됩니다. macOS에서는 Excel의 수동 기능을 사용해주세요.", err=True
                )
            else:
                typer.echo(
                    "💡 Excel이 설치되어 있는지 확인하고, 파일이 다른 프로그램에서 사용 중이지 않은지 확인하세요.", err=True
                )
        raise typer.Exit(1)

    except Exception as e:
        error_response = create_error_response(e, "pivot-delete")
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
