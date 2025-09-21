"""
차트 목록 조회 명령어
워크시트의 모든 차트 정보를 조회하는 기능
"""

import json
import platform
from pathlib import Path
from typing import Optional

import typer
import xlwings as xw

from pyhub_office_automation.version import get_version

from .utils import create_error_response, create_success_response, get_or_open_workbook, get_sheet, normalize_path


def get_chart_type_name(chart_obj):
    """차트 객체에서 차트 타입 이름을 추출"""
    try:
        if platform.system() == "Windows":
            # Windows에서는 API를 통해 정확한 차트 타입 가져오기
            chart_type_value = chart_obj.api.ChartType

            # 주요 차트 타입 매핑
            type_map = {
                51: "column_clustered",
                52: "column_stacked",
                53: "column_stacked_100",
                57: "bar_clustered",
                58: "bar_stacked",
                59: "bar_stacked_100",
                4: "line",
                65: "line_markers",
                5: "pie",
                -4120: "doughnut",
                1: "area",
                76: "area_stacked",
                77: "area_stacked_100",
                -4169: "scatter",
                74: "scatter_lines",
                72: "scatter_smooth",
                15: "bubble",
            }

            return type_map.get(chart_type_value, f"unknown_{chart_type_value}")
        else:
            # macOS에서는 기본값 반환
            return "chart"
    except:
        return "unknown"


def get_chart_title(chart_obj):
    """차트 제목 추출"""
    try:
        if hasattr(chart_obj, "api") and chart_obj.api.HasTitle:
            return chart_obj.api.ChartTitle.Text
        return None
    except:
        return None


def get_chart_legend_info(chart_obj):
    """범례 정보 추출"""
    try:
        if hasattr(chart_obj, "api"):
            has_legend = chart_obj.api.HasLegend
            if has_legend and platform.system() == "Windows":
                position_map = {-4107: "bottom", -4131: "corner", -4152: "left", -4161: "right", -4160: "top"}
                position = position_map.get(chart_obj.api.Legend.Position, "unknown")
                return {"has_legend": True, "position": position}
            return {"has_legend": has_legend, "position": None}
    except:
        return {"has_legend": False, "position": None}


def get_chart_data_source(chart_obj):
    """차트 데이터 소스 범위 추출"""
    try:
        if hasattr(chart_obj, "api") and platform.system() == "Windows":
            # Windows에서는 Series 데이터 소스 조회
            series_collection = chart_obj.api.FullSeriesCollection()
            if series_collection.Count > 0:
                first_series = series_collection(1)
                formula = first_series.Formula
                # 수식에서 범위 추출 (간단한 파싱)
                if "=" in formula and "!" in formula:
                    # =SERIES(,Sheet1!$A$1:$A$10,Sheet1!$B$1:$B$10,1) 형태에서 범위 추출
                    parts = formula.split(",")
                    if len(parts) >= 3:
                        range_part = parts[2].strip()
                        return range_part
            return None
    except:
        return None


def chart_list(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="차트를 조회할 Excel 파일의 절대 경로"),
    use_active: bool = typer.Option(False, "--use-active", help="현재 활성 워크북 사용"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help='열린 워크북 이름으로 접근 (예: "Sales.xlsx")'),
    sheet: Optional[str] = typer.Option(None, "--sheet", help="특정 시트의 차트만 조회 (지정하지 않으면 모든 시트)"),
    detailed: bool = typer.Option(False, "--detailed", help="차트의 상세 정보 포함"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택 (json/text)"),
    visible: bool = typer.Option(False, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부 (기본값: False)"),
):
    """
    워크시트의 모든 차트 정보를 조회합니다.

    워크북의 모든 시트를 검색하여 차트를 찾고, 각 차트의 기본 정보나 상세 정보를 반환합니다.
    차트 관리, 대시보드 분석, 차트 인벤토리 파악에 유용합니다.

    === 워크북 접근 방법 ===
    - --file-path: 파일 경로로 워크북 열기
    - --use-active: 현재 활성 워크북 사용 (가장 간편)
    - --workbook-name: 열린 워크북 이름으로 접근 (예: "Sales.xlsx")

    === 조회 범위 지정 ===
    --sheet 옵션으로 조회 범위를 제한할 수 있습니다:

    • 전체 워크북: 옵션 생략 (모든 시트의 차트 조회)
    • 특정 시트: --sheet "Dashboard" (해당 시트만 조회)
    • 여러 시트: 명령어를 여러 번 실행

    === 정보 상세도 선택 ===
    --detailed 플래그로 정보의 상세도를 조절합니다:

    ▶ 기본 정보 (--detailed 없음):
      • 차트 이름, 인덱스 번호
      • 위치 (셀 주소), 크기 (픽셀)
      • 소속 시트명

    ▶ 상세 정보 (--detailed 포함):
      • 기본 정보 + 추가 정보
      • 차트 유형 (column, pie, line 등)
      • 차트 제목, 범례 설정
      • 데이터 소스 범위 (Windows만)
      • 차트 스타일 정보

    === 활용 시나리오별 예제 ===

    # 1. 현재 워크북의 모든 차트 간단 조회
    oa excel chart-list --use-active

    # 2. 특정 시트의 차트만 상세 조회
    oa excel chart-list --use-active --sheet "Dashboard" --detailed

    # 3. 파일의 모든 차트 상세 분석
    oa excel chart-list --file-path "report.xlsx" --detailed

    # 4. 차트 인벤토리 텍스트 형식으로 출력
    oa excel chart-list --workbook-name "Sales.xlsx" --detailed --format text

    === 출력 활용 방법 ===
    • JSON 출력: AI 에이전트가 파싱하여 차트 정보 분석
    • TEXT 출력: 사람이 읽기 쉬운 형태로 차트 목록 확인
    • 차트 이름/인덱스: 다른 차트 명령어의 입력값으로 활용
    • 위치 정보: 차트 배치 현황 파악 및 재배치 계획
    • 데이터 소스: 차트 업데이트 및 수정 시 참고

    === 플랫폼별 차이점 ===
    • Windows: 모든 정보 제공 (차트 타입, 데이터 소스 등)
    • macOS: 기본 정보만 제공 (이름, 위치, 크기)
    """
    # 입력 값 검증
    if output_format not in ["json", "text"]:
        raise ValueError(f"잘못된 출력 형식: {output_format}. 사용 가능한 형식: json, text")

    book = None

    try:
        # 워크북 연결
        book = get_or_open_workbook(file_path=file_path, workbook_name=workbook_name, use_active=use_active, visible=visible)

        charts_info = []
        total_charts = 0

        # 시트 목록 결정
        if sheet:
            try:
                sheets_to_check = [get_sheet(book, sheet)]
            except ValueError:
                raise ValueError(f"시트 '{sheet}'를 찾을 수 없습니다")
        else:
            sheets_to_check = book.sheets

        # 각 시트의 차트 검색
        for worksheet in sheets_to_check:
            sheet_charts = []

            try:
                for i, chart in enumerate(worksheet.charts):
                    chart_info = {
                        "index": i,
                        "name": chart.name,
                        "sheet": worksheet.name,
                        "position": {"left": chart.left, "top": chart.top},
                        "dimensions": {"width": chart.width, "height": chart.height},
                    }

                    # 상세 정보 추가
                    if detailed:
                        # 차트 타입
                        chart_info["chart_type"] = get_chart_type_name(chart)

                        # 차트 제목
                        title = get_chart_title(chart)
                        if title:
                            chart_info["title"] = title

                        # 범례 정보
                        legend_info = get_chart_legend_info(chart)
                        chart_info["legend"] = legend_info

                        # 데이터 소스 (Windows에서만 정확히 가져올 수 있음)
                        data_source = get_chart_data_source(chart)
                        if data_source:
                            chart_info["data_source"] = data_source

                        # 플랫폼별 추가 정보
                        chart_info["platform_support"] = {
                            "current_platform": platform.system(),
                            "full_features_available": platform.system() == "Windows",
                        }

                    sheet_charts.append(chart_info)
                    total_charts += 1

            except Exception as e:
                # 특정 시트에서 차트 조회 실패해도 계속 진행
                sheet_charts.append({"error": f"시트 '{worksheet.name}'에서 차트 조회 실패: {str(e)}"})

            if sheet_charts:
                charts_info.extend(sheet_charts)

        # 응답 데이터 구성
        response_data = {
            "workbook": book.name,
            "total_charts": total_charts,
            "charts": charts_info,
            "query_info": {
                "target_sheet": sheet if sheet else "all_sheets",
                "detailed": detailed,
                "platform": platform.system(),
            },
        }

        if sheet:
            response_data["sheet"] = sheet
        else:
            response_data["sheets_checked"] = [ws.name for ws in sheets_to_check]

        response = create_success_response(
            data=response_data, command="chart-list", message=f"{total_charts}개의 차트를 찾았습니다"
        )

        if output_format == "json":
            print(json.dumps(response, ensure_ascii=False, indent=2))
        else:
            # 텍스트 형식 출력
            print(f"=== 차트 목록 ===")
            print(f"워크북: {book.name}")
            print(f"총 차트 수: {total_charts}")
            print()

            if total_charts == 0:
                print("차트가 없습니다.")
            else:
                for chart in charts_info:
                    if "error" in chart:
                        print(f"❌ {chart['error']}")
                        continue

                    print(f"📊 {chart['name']}")
                    print(f"   시트: {chart['sheet']}")
                    print(f"   위치: ({chart['position']['left']}, {chart['position']['top']})")
                    print(f"   크기: {chart['dimensions']['width']} x {chart['dimensions']['height']}")

                    if detailed:
                        print(f"   타입: {chart.get('chart_type', 'unknown')}")
                        if chart.get("title"):
                            print(f"   제목: {chart['title']}")
                        if chart.get("legend"):
                            legend = chart["legend"]
                            if legend["has_legend"]:
                                print(f"   범례: {legend.get('position', '위치 불명')}")
                            else:
                                print(f"   범례: 없음")
                        if chart.get("data_source"):
                            print(f"   데이터: {chart['data_source']}")
                    print()

    except Exception as e:
        error_response = create_error_response(e, "chart-list")
        if output_format == "json":
            print(json.dumps(error_response, ensure_ascii=False, indent=2))
        else:
            print(f"오류: {str(e)}")
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
    chart_list()
