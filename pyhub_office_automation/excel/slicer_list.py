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

from .utils import (
    ExecutionTimer,
    create_error_response,
    create_success_response,
    get_or_open_workbook,
    get_slicers_info,
    normalize_path,
)


def slicer_list(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="슬라이서를 조회할 Excel 파일의 절대 경로"),
    use_active: bool = typer.Option(False, "--use-active", help="현재 활성 워크북 사용"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help='열린 워크북 이름으로 접근 (예: "Sales.xlsx")'),
    detailed: bool = typer.Option(False, "--detailed", help="상세 정보 포함 (슬라이서 항목, 연결된 피벗테이블 등)"),
    include_items: bool = typer.Option(False, "--include-items", help="슬라이서 항목 목록 포함"),
    show_connections: bool = typer.Option(False, "--show-connections", help="연결된 피벗테이블 정보 표시"),
    filter_field: Optional[str] = typer.Option(None, "--filter-field", help="특정 필드의 슬라이서만 필터링"),
    filter_sheet: Optional[str] = typer.Option(None, "--filter-sheet", help="특정 시트의 슬라이서만 필터링"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택 (json/text)"),
    visible: bool = typer.Option(False, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부 (기본값: False)"),
):
    """
    Excel 워크북의 모든 슬라이서 정보를 조회합니다.

    슬라이서의 기본 정보부터 상세 설정, 연결된 피벗테이블, 현재 선택 상태까지
    조회할 수 있으며, 대시보드 분석 및 슬라이서 관리에 유용합니다.

    === 워크북 접근 방법 ===
    - --file-path: 파일 경로로 워크북 열기
    - --use-active: 현재 활성 워크북 사용
    - --workbook-name: 열린 워크북 이름으로 접근 (예: "Sales.xlsx")

    === 조회 옵션 ===
    • 기본 조회: 슬라이서 이름, 위치, 크기, 필드명
    • --detailed: 스타일, 레이아웃 설정 등 상세 정보
    • --include-items: 슬라이서 항목 목록과 선택 상태
    • --show-connections: 연결된 피벗테이블 정보

    === 필터링 옵션 ===
    • --filter-field: 특정 필드의 슬라이서만 조회
    • --filter-sheet: 특정 시트의 슬라이서만 조회

    === 기본 정보 항목 ===
    • name: 슬라이서 이름
    • field_name: 기반 필드명
    • position: {left, top} 위치 정보
    • size: {width, height} 크기 정보
    • sheet: 배치된 시트명

    === 상세 정보 항목 (--detailed) ===
    • source_name: 원본 데이터 소스
    • slicer_items: 항목 목록 (이름, 선택 상태)
    • connected_pivot_tables: 연결된 피벗테이블 목록
    • style_settings: 스타일 및 레이아웃 설정

    === 사용 시나리오 ===

    # 1. 워크북의 모든 슬라이서 기본 정보 조회
    oa excel slicer-list --use-active

    # 2. 상세한 슬라이서 정보 조회
    oa excel slicer-list --file-path "dashboard.xlsx" --detailed

    # 3. 슬라이서 항목과 선택 상태까지 포함한 전체 정보
    oa excel slicer-list --use-active --detailed --include-items --show-connections

    # 4. 특정 필드의 슬라이서만 조회
    oa excel slicer-list --use-active --filter-field "지역" --detailed

    # 5. 특정 시트의 슬라이서만 조회
    oa excel slicer-list --use-active --filter-sheet "Dashboard" --include-items

    # 6. 대시보드 슬라이서 현황 분석
    oa excel slicer-list --workbook-name "SalesReport.xlsx" \\
        --detailed --include-items --show-connections --format json

    === 출력 예제 ===
    ```json
    {
      "success": true,
      "data": {
        "slicers": [
          {
            "name": "RegionSlicer",
            "field_name": "지역",
            "position": {"left": 100, "top": 400},
            "size": {"width": 200, "height": 120},
            "sheet": "Dashboard",
            "connected_pivot_tables": ["SalesPivot", "TrendPivot"],
            "slicer_items": [
              {"name": "서울", "selected": true},
              {"name": "부산", "selected": false}
            ]
          }
        ],
        "total_slicers": 1
      }
    }
    ```

    === 대시보드 관리 활용 ===
    • 슬라이서 배치 현황 파악
    • 필터 연결 상태 확인
    • 슬라이서 간 겹침 검사
    • 선택 상태 모니터링
    • 대시보드 구조 분석

    === 연결 상태 분석 ===
    ```bash
    # 연결이 끊어진 슬라이서 찾기
    oa excel slicer-list --use-active --show-connections | \\
        grep -A 5 '"connected_pivot_tables": \\[\\]'

    # 특정 피벗테이블에 연결된 모든 슬라이서 확인
    oa excel slicer-list --use-active --show-connections | \\
        grep -B 5 -A 5 "SalesPivot"
    ```

    === 문제 해결 가이드 ===
    • 슬라이서가 표시되지 않는 경우: 위치 정보 확인
    • 필터가 작동하지 않는 경우: 연결된 피벗테이블 확인
    • 성능이 느린 경우: 과도한 슬라이서 항목 확인
    • 레이아웃이 깨진 경우: 크기 및 배치 정보 확인

    === 주의사항 ===
    • Windows에서만 완전한 정보 제공
    • macOS에서는 기본 정보만 제한적 지원
    • 대용량 데이터의 경우 조회 시간이 오래 걸릴 수 있음
    • 슬라이서 항목이 많은 경우 --include-items 주의
    """
    book = None

    try:
        with ExecutionTimer() as timer:
            # 워크북 연결
            book = get_or_open_workbook(
                file_path=file_path, workbook_name=workbook_name, use_active=use_active, visible=visible
            )

            # 슬라이서 정보 수집
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


if __name__ == "__main__":
    slicer_list()
