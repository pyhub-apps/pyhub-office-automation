"""
슬라이서 목록 조회 명령어
xlwings를 활용한 Excel 슬라이서 정보 수집 기능
"""

import json
import platform
from typing import Optional

import typer
import xlwings as xw

from pyhub_office_automation.version import get_version

from .engines import get_engine
from .utils import (
    ExecutionTimer,
    analyze_slicer_conflicts,
    create_error_response,
    create_success_response,
    get_or_open_workbook,
    get_slicers_info,
    normalize_path,
)


def slicer_list(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="슬라이서를 조회할 Excel 파일의 절대 경로"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help='열린 워크북 이름으로 접근 (예: "Sales.xlsx")'),
    brief: bool = typer.Option(False, "--brief", help="간단한 정보만 포함 (기본: 상세 정보 포함)"),
    detailed: bool = typer.Option(
        True, "--detailed/--no-detailed", help="상세 정보 포함 (슬라이서 항목, 연결된 피벗테이블 등)"
    ),
    include_items: bool = typer.Option(True, "--include-items/--no-include-items", help="슬라이서 항목 목록 포함"),
    show_connections: bool = typer.Option(
        True, "--show-connections/--no-show-connections", help="연결된 피벗테이블 정보 표시"
    ),
    show_conflicts: bool = typer.Option(False, "--show-conflicts", help="SlicerCache 충돌 가능성 분석 표시 (Issue #71)"),
    filter_field: Optional[str] = typer.Option(None, "--filter-field", help="특정 필드의 슬라이서만 필터링"),
    filter_sheet: Optional[str] = typer.Option(None, "--filter-sheet", help="특정 시트의 슬라이서만 필터링"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택 (json/text)"),
    visible: bool = typer.Option(False, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부 (기본값: False)"),
):
    """
    🔍 Excel 워크북의 모든 슬라이서 정보를 조회합니다.

    슬라이서의 기본 정보부터 상세 설정, 연결된 피벗테이블, 현재 선택 상태까지
    조회할 수 있으며, 대시보드 분석 및 슬라이서 관리에 유용합니다.

    ## 📁 워크북 접근 방법

    - `--file-path`: 파일 경로로 워크북 열기
    - `--workbook-name`: 열린 워크북 이름으로 접근 (예: "Sales.xlsx")

    ## 📊 조회 옵션

    - `--detailed`: 스타일, 레이아웃 설정 등 상세 정보
    - `--include-items`: 슬라이서 항목 목록과 선택 상태
    - `--show-connections`: 연결된 피벗테이블 정보
    - `--show-conflicts`: SlicerCache 충돌 가능성 분석 (Issue #71)

    ## 🔍 필터링 옵션

    - `--filter-field`: 특정 필드의 슬라이서만 조회
    - `--filter-sheet`: 특정 시트의 슬라이서만 조회

    ## 🚀 사용 예시

    **기본 조회:**
    ```bash
    oa excel slicer-list
    ```

    **상세 정보:**
    ```bash
    oa excel slicer-list --detailed --include-items --show-connections
    ```

    **충돌 분석:**
    ```bash
    oa excel slicer-list --show-conflicts
    ```

    **필터링:**
    ```bash
    oa excel slicer-list --filter-field "지역" --detailed
    ```

    ## ⚠️ 주의사항

    - Windows에서만 완전한 정보 제공
    - macOS에서는 기본 정보만 제한적 지원
    - 대용량 데이터의 경우 조회 시간이 오래 걸릴 수 있음
    """
    book = None

    try:
        with ExecutionTimer() as timer:
            # brief 옵션 처리 - 간단한 정보만 포함
            if brief:
                detailed = False
                include_items = False
                show_connections = False

            # 워크북 연결
            book = get_or_open_workbook(file_path=file_path, workbook_name=workbook_name, visible=visible)

            # Engine 가져오기
            engine = get_engine()

            # 슬라이서 정보 수집 (Engine Layer 사용)
            try:
                slicers_info = engine.list_slicers(workbook=book.api)
            except Exception as e:
                # Fallback to utility function if engine method fails
                slicers_info = get_slicers_info(book)

            # 필터링 적용
            if filter_field:
                filtered_slicers = []
                for slicer_info in slicers_info:
                    if filter_field.lower() in slicer_info.get("field_name", "").lower():
                        filtered_slicers.append(slicer_info)
                slicers_info = filtered_slicers

            if filter_sheet:
                filtered_slicers = []
                for slicer_info in slicers_info:
                    if filter_sheet.lower() in slicer_info.get("sheet", "").lower():
                        filtered_slicers.append(slicer_info)
                slicers_info = filtered_slicers

            # 상세 정보 처리
            if not detailed:
                # 기본 정보만 포함
                for slicer_info in slicers_info:
                    # 불필요한 정보 제거
                    simplified_info = {
                        "name": slicer_info.get("name"),
                        "field_name": slicer_info.get("field_name"),
                        "position": slicer_info.get("position"),
                        "size": slicer_info.get("size"),
                        "sheet": slicer_info.get("sheet"),
                    }

                    # 기본 연결 정보는 유지
                    if slicer_info.get("connected_pivot_tables"):
                        simplified_info["connected_pivot_tables"] = len(slicer_info["connected_pivot_tables"])

                    # 원본 정보 교체
                    for key in list(slicer_info.keys()):
                        del slicer_info[key]
                    slicer_info.update(simplified_info)

            # 선택적 정보 제거
            if not include_items:
                for slicer_info in slicers_info:
                    if "slicer_items" in slicer_info:
                        # 항목 개수만 유지
                        item_count = len(slicer_info["slicer_items"])
                        selected_count = sum(1 for item in slicer_info["slicer_items"] if item.get("selected", False))
                        del slicer_info["slicer_items"]
                        slicer_info["item_summary"] = {"total_items": item_count, "selected_items": selected_count}

            if not show_connections:
                for slicer_info in slicers_info:
                    if "connected_pivot_tables" in slicer_info:
                        # 연결 개수만 유지
                        connection_count = len(slicer_info["connected_pivot_tables"])
                        del slicer_info["connected_pivot_tables"]
                        if detailed:
                            slicer_info["connection_count"] = connection_count

            # Windows에서 추가 정보 수집 (detailed 모드)
            if detailed and platform.system() == "Windows":
                for slicer_info in slicers_info:
                    try:
                        # 추가 슬라이서 설정 정보 수집
                        slicer_info["platform_info"] = {"full_support": True, "additional_settings_available": True}
                    except Exception:
                        pass

            # 응답 데이터 구성
            response_data = {
                "slicers": slicers_info,
                "total_slicers": len(slicers_info),
                "workbook": normalize_path(book.name),
                "query_options": {
                    "detailed": detailed,
                    "include_items": include_items,
                    "show_connections": show_connections,
                    "filter_field": filter_field,
                    "filter_sheet": filter_sheet,
                },
            }

            # 플랫폼별 지원 정보
            if platform.system() != "Windows":
                response_data["platform_note"] = "macOS에서는 제한된 슬라이서 정보만 제공됩니다"

            # 통계 정보
            if slicers_info:
                # 필드별 통계
                field_stats = {}
                sheet_stats = {}
                total_items = 0
                total_selected = 0

                for slicer_info in slicers_info:
                    field_name = slicer_info.get("field_name", "Unknown")
                    sheet_name = slicer_info.get("sheet", "Unknown")

                    field_stats[field_name] = field_stats.get(field_name, 0) + 1
                    sheet_stats[sheet_name] = sheet_stats.get(sheet_name, 0) + 1

                    # 항목 통계
                    if "item_summary" in slicer_info:
                        total_items += slicer_info["item_summary"]["total_items"]
                        total_selected += slicer_info["item_summary"]["selected_items"]
                    elif "slicer_items" in slicer_info:
                        total_items += len(slicer_info["slicer_items"])
                        total_selected += sum(1 for item in slicer_info["slicer_items"] if item.get("selected", False))

                response_data["statistics"] = {
                    "slicers_by_field": field_stats,
                    "slicers_by_sheet": sheet_stats,
                    "total_slicer_items": total_items,
                    "total_selected_items": total_selected,
                }

            # SlicerCache 충돌 분석 (Issue #71)
            if show_conflicts and platform.system() == "Windows":
                conflict_analysis = analyze_slicer_conflicts(slicers_info)
                response_data["conflict_analysis"] = conflict_analysis

            message = f"{len(slicers_info)}개의 슬라이서 정보를 조회했습니다"
            if filter_field or filter_sheet:
                message += " (필터 적용됨)"

            response = create_success_response(
                data=response_data,
                command="slicer-list",
                message=message,
                execution_time_ms=timer.execution_time_ms,
                book=book,
                slicers_count=len(slicers_info),
            )

            print(json.dumps(response, ensure_ascii=False, indent=2))

    except Exception as e:
        error_response = create_error_response(e, "slicer-list")
        print(json.dumps(error_response, ensure_ascii=False, indent=2))
        return 1

    finally:
        # 새로 생성한 워크북인 경우에만 정리
        if book and file_path and not workbook_name:
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


if __name__ == "__main__":
    slicer_list()
