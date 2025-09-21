"""
슬라이서 연결 관리 명령어
xlwings를 활용한 Excel 슬라이서와 피벗테이블 연결 기능
"""

import json
import platform
from typing import Optional

import typer
import xlwings as xw

from pyhub_office_automation.version import get_version

from .utils import (
    ExecutionTimer,
    create_error_response,
    create_success_response,
    get_or_open_workbook,
    get_pivot_tables,
    get_sheet,
    get_slicer_by_name,
    normalize_path,
)


def slicer_connect(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="슬라이서 연결을 관리할 Excel 파일의 절대 경로"),
    use_active: bool = typer.Option(False, "--use-active", help="현재 활성 워크북 사용"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help='열린 워크북 이름으로 접근 (예: "Sales.xlsx")'),
    slicer_name: str = typer.Option(..., "--slicer-name", help="연결을 관리할 슬라이서 이름"),
    action: str = typer.Option(..., "--action", help="작업 유형: connect(연결), disconnect(연결 해제), list(연결 상태 조회)"),
    pivot_tables: Optional[str] = typer.Option(
        None, "--pivot-tables", help='연결할 피벗테이블 이름들 (쉼표로 구분, 예: "Pivot1,Pivot2")'
    ),
    all_pivots: bool = typer.Option(False, "--all-pivots", help="워크북의 모든 피벗테이블에 연결/해제"),
    target_sheet: Optional[str] = typer.Option(None, "--target-sheet", help="특정 시트의 피벗테이블만 대상으로 지정"),
    force: bool = typer.Option(False, "--force", help="강제 연결/해제 (호환성 문제 무시)"),
    dry_run: bool = typer.Option(False, "--dry-run", help="실제 연결하지 않고 대상만 확인"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택 (json/text)"),
    visible: bool = typer.Option(False, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부 (기본값: False)"),
    save: bool = typer.Option(True, "--save", help="연결 변경 후 파일 저장 여부 (기본값: True)"),
):
    """
    Excel 슬라이서와 피벗테이블의 연결을 관리합니다.

    하나의 슬라이서를 여러 피벗테이블에 연결하여 통합 필터링 기능을 구현하거나,
    불필요한 연결을 해제하여 성능을 최적화할 수 있습니다.

    === 워크북 접근 방법 ===
    - --file-path: 파일 경로로 워크북 열기
    - --use-active: 현재 활성 워크북 사용
    - --workbook-name: 열린 워크북 이름으로 접근 (예: "Sales.xlsx")

    === 작업 유형 ===
    • connect: 슬라이서를 피벗테이블에 연결
    • disconnect: 슬라이서와 피벗테이블 연결 해제
    • list: 현재 연결 상태 조회

    === 연결 대상 지정 ===
    • --pivot-tables: 특정 피벗테이블들 (쉼표 구분)
    • --all-pivots: 워크북의 모든 피벗테이블
    • --target-sheet: 특정 시트의 피벗테이블만

    === 안전 기능 ===
    • --dry-run: 실제 연결 전 대상 확인
    • --force: 호환성 경고 무시하고 강제 실행
    • 연결 전 호환성 검사 (필드명, 데이터 소스)

    === 슬라이서 연결 시나리오 ===

    # 1. 지역 슬라이서를 여러 피벗테이블에 연결 (통합 필터링)
    oa excel slicer-connect --use-active --slicer-name "RegionSlicer" \\
        --action connect --pivot-tables "SalesPivot,TrendPivot,KPIPivot"

    # 2. 모든 피벗테이블에 월별 슬라이서 연결
    oa excel slicer-connect --use-active --slicer-name "MonthSlicer" \\
        --action connect --all-pivots

    # 3. 특정 시트의 피벗테이블에만 연결
    oa excel slicer-connect --use-active --slicer-name "CategorySlicer" \\
        --action connect --target-sheet "Analysis" --all-pivots

    # 4. 연결 전 호환성 확인 (dry-run)
    oa excel slicer-connect --use-active --slicer-name "ProductSlicer" \\
        --action connect --pivot-tables "DetailPivot,SummaryPivot" --dry-run

    === 연결 해제 시나리오 ===

    # 1. 특정 피벗테이블과의 연결 해제
    oa excel slicer-connect --use-active --slicer-name "RegionSlicer" \\
        --action disconnect --pivot-tables "OldPivot"

    # 2. 모든 연결 해제 (슬라이서 독립화)
    oa excel slicer-connect --use-active --slicer-name "TempSlicer" \\
        --action disconnect --all-pivots

    # 3. 강제 연결 해제 (오류 무시)
    oa excel slicer-connect --use-active --slicer-name "BrokenSlicer" \\
        --action disconnect --all-pivots --force

    === 연결 상태 조회 ===

    # 1. 슬라이서의 모든 연결 상태 확인
    oa excel slicer-connect --use-active --slicer-name "RegionSlicer" --action list

    # 2. 연결 상태와 호환성 정보 확인
    oa excel slicer-connect --file-path "dashboard.xlsx" \\
        --slicer-name "CategorySlicer" --action list --format json

    === 대시보드 통합 필터링 구성 ===

    # 단계별 대시보드 슬라이서 연결 구성
    # 1. 지역 필터 - 모든 분석에 적용
    oa excel slicer-connect --use-active --slicer-name "RegionSlicer" \\
        --action connect --pivot-tables "MainSales,TrendAnalysis,KPIDashboard,DetailReport"

    # 2. 기간 필터 - 시계열 분석에만 적용
    oa excel slicer-connect --use-active --slicer-name "DateSlicer" \\
        --action connect --pivot-tables "TrendAnalysis,MonthlyReport"

    # 3. 제품 분류 - 제품 관련 분석에만 적용
    oa excel slicer-connect --use-active --slicer-name "CategorySlicer" \\
        --action connect --pivot-tables "ProductAnalysis,CategorySales"

    # 4. 전체 연결 상태 확인
    oa excel slicer-connect --use-active --slicer-name "RegionSlicer" --action list
    oa excel slicer-connect --use-active --slicer-name "DateSlicer" --action list
    oa excel slicer-connect --use-active --slicer-name "CategorySlicer" --action list

    === 고급 연결 관리 ===

    # 성능 최적화를 위한 선별적 연결
    # 1. 메인 대시보드 슬라이서는 핵심 피벗테이블에만 연결
    oa excel slicer-connect --use-active --slicer-name "MainRegionSlicer" \\
        --action connect --pivot-tables "PrimarySales,CoreKPI"

    # 2. 상세 분석 슬라이서는 해당 분석용 피벗테이블에만 연결
    oa excel slicer-connect --use-active --slicer-name "DetailCategorySlicer" \\
        --action connect --target-sheet "DetailAnalysis" --all-pivots

    # 3. 임시 슬라이서 정리 (모든 연결 해제)
    oa excel slicer-connect --use-active --slicer-name "TempFilterSlicer" \\
        --action disconnect --all-pivots

    === 연결 호환성 조건 ===
    • 동일한 필드명을 가진 피벗테이블만 연결 가능
    • 동일한 데이터 소스 기반 피벗테이블 권장
    • 필드 데이터 타입이 일치해야 함
    • OLAP 큐브와 테이블 데이터 간 연결 제한

    === 문제 해결 가이드 ===
    • "호환되지 않는 피벗테이블" 오류: 필드명과 데이터 소스 확인
    • "연결 실패" 오류: --force 옵션 사용 또는 수동 재생성
    • 성능 저하: 불필요한 연결 해제로 최적화
    • 필터 동작 안 함: 연결 상태 확인 후 재연결

    === 주의사항 ===
    • Windows에서만 완전 지원
    • 연결 변경 시 기존 필터 선택 상태 초기화 가능
    • 대용량 데이터의 경우 연결 작업 시간 소요
    • 순환 참조나 충돌 방지를 위한 신중한 연결 계획 필요
    """
    book = None

    try:
        # action 검증
        if action not in ["connect", "disconnect", "list"]:
            raise ValueError(f"action은 'connect', 'disconnect', 'list' 중 하나여야 합니다. 입력된 값: {action}")

        with ExecutionTimer() as timer:
            # Windows 플랫폼 확인
            if platform.system() != "Windows":
                raise RuntimeError("슬라이서 연결 관리는 Windows에서만 지원됩니다")

            # 워크북 연결
            book = get_or_open_workbook(
                file_path=file_path, workbook_name=workbook_name, use_active=use_active, visible=visible
            )

            # 슬라이서 찾기
            slicer_cache = get_slicer_by_name(book, slicer_name)
            if not slicer_cache:
                raise ValueError(f"슬라이서 '{slicer_name}'을 찾을 수 없습니다")

            # 작업별 처리
            if action == "list":
                result = handle_list_connections(book, slicer_cache, slicer_name)
            elif action == "connect":
                result = handle_connect_action(
                    book, slicer_cache, slicer_name, pivot_tables, all_pivots, target_sheet, force, dry_run
                )
            else:  # disconnect
                result = handle_disconnect_action(
                    book, slicer_cache, slicer_name, pivot_tables, all_pivots, target_sheet, force, dry_run
                )

            # 파일 저장
            if save and file_path and not dry_run and action != "list":
                book.save()

            # 성공 응답 생성
            response_data = {
                "slicer_name": slicer_name,
                "action": action,
                "dry_run": dry_run,
                "workbook": normalize_path(book.name),
                **result,
            }

            message = result.get("message", f"{action} 작업이 완료되었습니다")
            if dry_run:
                message = f"[DRY RUN] {message}"

            response = create_success_response(
                data=response_data,
                command="slicer-connect",
                message=message,
                execution_time_ms=timer.execution_time_ms,
                book=book,
                action=action,
                affected_connections=result.get("affected_connections", 0),
            )

            print(json.dumps(response, ensure_ascii=False, indent=2))

    except Exception as e:
        error_response = create_error_response(e, "slicer-connect")
        print(json.dumps(error_response, ensure_ascii=False, indent=2))
        return 1

    finally:
        # 새로 생성한 워크북인 경우에만 정리
        if book and file_path and not use_active and not workbook_name:
            try:
                if visible:
                    # 화면에 표시하는 경우 닫지 않음
                    pass
                else:
                    # 백그라운드 실행인 경우 앱 정리
                    book.app.quit()
            except:
                pass

    return 0


def handle_list_connections(book, slicer_cache, slicer_name):
    """연결 상태 조회"""
    try:
        current_connections = []

        # 현재 연결된 피벗테이블 목록
        for pivot_table in slicer_cache.PivotTables():
            connection_info = {
                "pivot_table_name": pivot_table.Name,
                "sheet": pivot_table.Parent.Name,
                "location": pivot_table.TableRange1.Address,
                "source_data": getattr(pivot_table, "SourceData", "Unknown"),
            }

            # 호환성 확인
            try:
                # 슬라이서 필드와 피벗테이블 필드 호환성 검사
                slicer_field = slicer_cache.SourceName
                pivot_fields = [field.Name for field in pivot_table.PivotFields()]

                connection_info["field_compatible"] = slicer_field in pivot_fields
                connection_info["slicer_field"] = slicer_field

            except Exception:
                connection_info["field_compatible"] = "unknown"

            current_connections.append(connection_info)

        # 슬라이서 기본 정보
        slicer_info = {
            "name": slicer_name,
            "source_name": slicer_cache.SourceName,
            "field_name": getattr(slicer_cache, "SourceName", "Unknown"),
        }

        return {
            "slicer_info": slicer_info,
            "current_connections": current_connections,
            "total_connections": len(current_connections),
            "message": f"슬라이서 '{slicer_name}'은 {len(current_connections)}개의 피벗테이블에 연결되어 있습니다",
        }

    except Exception as e:
        raise RuntimeError(f"연결 상태 조회 실패: {str(e)}")


def handle_connect_action(book, slicer_cache, slicer_name, pivot_tables, all_pivots, target_sheet, force, dry_run):
    """연결 작업 처리"""
    if not pivot_tables and not all_pivots:
        raise ValueError("연결할 피벗테이블을 지정해야 합니다 (--pivot-tables 또는 --all-pivots)")

    # 대상 피벗테이블 수집
    target_pivot_tables = collect_target_pivot_tables(book, pivot_tables, all_pivots, target_sheet)

    if not target_pivot_tables:
        raise ValueError("연결할 피벗테이블을 찾을 수 없습니다")

    # 호환성 검사
    slicer_field = slicer_cache.SourceName
    compatible_pivots = []
    incompatible_pivots = []

    for pivot_info in target_pivot_tables:
        try:
            pivot_obj = pivot_info["pivot_object"]
            pivot_fields = [field.Name for field in pivot_obj.PivotFields()]

            if slicer_field in pivot_fields:
                compatible_pivots.append(pivot_info)
            else:
                incompatible_pivots.append(
                    {"name": pivot_info["name"], "reason": f"필드 '{slicer_field}' 없음", "available_fields": pivot_fields}
                )
        except Exception as e:
            incompatible_pivots.append({"name": pivot_info["name"], "reason": f"검사 실패: {str(e)}"})

    if incompatible_pivots and not force:
        raise ValueError(
            f"호환되지 않는 피벗테이블이 있습니다: {[p['name'] for p in incompatible_pivots]}. "
            f"--force 옵션으로 강제 연결하거나 호환되는 피벗테이블만 지정하세요."
        )

    if dry_run:
        return {
            "target_pivot_tables": [p["name"] for p in compatible_pivots],
            "compatible_count": len(compatible_pivots),
            "incompatible_pivots": incompatible_pivots,
            "slicer_field": slicer_field,
            "affected_connections": len(compatible_pivots),
            "message": f"{len(compatible_pivots)}개의 피벗테이블에 연결될 예정입니다",
        }

    # 실제 연결 수행
    successful_connections = []
    failed_connections = []

    for pivot_info in compatible_pivots:
        try:
            # 슬라이서 캐시에 피벗테이블 추가
            slicer_cache.PivotTables.AddPivotTable(pivot_info["pivot_object"])
            successful_connections.append(pivot_info["name"])
        except Exception as e:
            failed_connections.append({"name": pivot_info["name"], "error": str(e)})

    # 강제 모드에서 호환되지 않는 피벗테이블도 시도
    if force and incompatible_pivots:
        for pivot_info in target_pivot_tables:
            if pivot_info["name"] in [p["name"] for p in incompatible_pivots]:
                try:
                    slicer_cache.PivotTables.AddPivotTable(pivot_info["pivot_object"])
                    successful_connections.append(pivot_info["name"] + " (강제)")
                except Exception as e:
                    failed_connections.append({"name": pivot_info["name"] + " (강제)", "error": str(e)})

    result = {
        "successful_connections": successful_connections,
        "total_connected": len(successful_connections),
        "affected_connections": len(successful_connections),
        "slicer_field": slicer_field,
    }

    if failed_connections:
        result["failed_connections"] = failed_connections

    if incompatible_pivots and not force:
        result["skipped_incompatible"] = incompatible_pivots

    message = f"{len(successful_connections)}개의 피벗테이블에 성공적으로 연결되었습니다"
    if failed_connections:
        message += f" ({len(failed_connections)}개 실패)"

    result["message"] = message
    return result


def handle_disconnect_action(book, slicer_cache, slicer_name, pivot_tables, all_pivots, target_sheet, force, dry_run):
    """연결 해제 작업 처리"""
    # 현재 연결된 피벗테이블 목록
    current_connections = []
    try:
        for pivot_table in slicer_cache.PivotTables():
            current_connections.append(
                {"name": pivot_table.Name, "sheet": pivot_table.Parent.Name, "pivot_object": pivot_table}
            )
    except Exception as e:
        if not force:
            raise RuntimeError(f"현재 연결 상태 확인 실패: {str(e)}")

    if not current_connections and not force:
        raise ValueError(f"슬라이서 '{slicer_name}'에 연결된 피벗테이블이 없습니다")

    # 해제 대상 결정
    if all_pivots:
        targets_to_disconnect = current_connections
    elif pivot_tables:
        target_names = [name.strip() for name in pivot_tables.split(",")]
        targets_to_disconnect = [conn for conn in current_connections if conn["name"] in target_names]

        missing_targets = [name for name in target_names if name not in [conn["name"] for conn in current_connections]]

        if missing_targets and not force:
            raise ValueError(f"다음 피벗테이블은 연결되어 있지 않습니다: {missing_targets}")
    else:
        raise ValueError("해제할 대상을 지정해야 합니다 (--pivot-tables 또는 --all-pivots)")

    # 시트 필터링
    if target_sheet:
        targets_to_disconnect = [conn for conn in targets_to_disconnect if conn["sheet"] == target_sheet]

    if not targets_to_disconnect:
        raise ValueError("해제할 대상이 없습니다")

    if dry_run:
        return {
            "targets_to_disconnect": [t["name"] for t in targets_to_disconnect],
            "total_to_disconnect": len(targets_to_disconnect),
            "current_connections": [c["name"] for c in current_connections],
            "affected_connections": len(targets_to_disconnect),
            "message": f"{len(targets_to_disconnect)}개의 연결이 해제될 예정입니다",
        }

    # 실제 해제 수행
    successful_disconnections = []
    failed_disconnections = []

    for target in targets_to_disconnect:
        try:
            # 특정 피벗테이블과의 연결 해제
            # Excel COM API를 통해 해제
            target["pivot_object"].PivotCache.RemoveSlicerCaches(slicer_cache)
            successful_disconnections.append(target["name"])
        except Exception as e:
            failed_disconnections.append({"name": target["name"], "error": str(e)})

    result = {
        "successful_disconnections": successful_disconnections,
        "total_disconnected": len(successful_disconnections),
        "affected_connections": len(successful_disconnections),
    }

    if failed_disconnections:
        result["failed_disconnections"] = failed_disconnections

    message = f"{len(successful_disconnections)}개의 연결이 성공적으로 해제되었습니다"
    if failed_disconnections:
        message += f" ({len(failed_disconnections)}개 실패)"

    result["message"] = message
    return result


def collect_target_pivot_tables(book, pivot_tables, all_pivots, target_sheet):
    """대상 피벗테이블 수집"""
    target_pivot_tables = []

    if all_pivots:
        # 모든 시트의 모든 피벗테이블
        for sheet in book.sheets:
            if target_sheet and sheet.name != target_sheet:
                continue

            pivot_tables_info = get_pivot_tables(sheet)
            for pt_info in pivot_tables_info:
                try:
                    # 실제 피벗테이블 객체 가져오기
                    for pt in sheet.api.PivotTables():
                        if pt.Name == pt_info["name"]:
                            target_pivot_tables.append({"name": pt_info["name"], "sheet": sheet.name, "pivot_object": pt})
                            break
                except Exception:
                    continue
    else:
        # 지정된 피벗테이블들
        target_names = [name.strip() for name in pivot_tables.split(",")]

        for sheet in book.sheets:
            if target_sheet and sheet.name != target_sheet:
                continue

            try:
                for pt in sheet.api.PivotTables():
                    if pt.Name in target_names:
                        target_pivot_tables.append({"name": pt.Name, "sheet": sheet.name, "pivot_object": pt})
            except Exception:
                continue

    return target_pivot_tables


if __name__ == "__main__":
    slicer_connect()
