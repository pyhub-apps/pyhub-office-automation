"""
차트 목록 조회 명령어
워크시트의 모든 차트 정보를 조회하는 기능
"""

import json
import platform
from pathlib import Path
from typing import Optional

import typer

from pyhub_office_automation.version import get_version

from .engines import get_engine
from .utils import create_error_response, create_success_response


def get_chart_legend_info(chart_obj):
    """
    Windows COM API를 사용하여 범례 정보를 추출합니다.
    detailed 옵션에서만 사용됩니다.
    """
    try:
        has_legend = chart_obj.HasLegend
        if has_legend and platform.system() == "Windows":
            position_map = {-4107: "bottom", -4131: "corner", -4152: "left", -4161: "right", -4160: "top"}
            position = position_map.get(chart_obj.Legend.Position, "unknown")
            return {"has_legend": True, "position": position}
        return {"has_legend": has_legend, "position": None}
    except:
        return {"has_legend": False, "position": None}


def chart_list(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="차트를 조회할 Excel 파일의 절대 경로"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help='열린 워크북 이름으로 접근 (예: "Sales.xlsx")'),
    sheet: Optional[str] = typer.Option(None, "--sheet", help="특정 시트의 차트만 조회 (지정하지 않으면 모든 시트)"),
    brief: bool = typer.Option(False, "--brief", help="간단한 정보만 포함 (기본: 상세 정보 포함)"),
    detailed: bool = typer.Option(True, "--detailed/--no-detailed", help="차트의 상세 정보 포함 (기본: True)"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택 (json/text)"),
    visible: bool = typer.Option(False, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부 (기본값: False)"),
):
    """
    워크시트의 모든 차트 정보를 조회합니다. (기본적으로 상세 정보 포함)

    워크북의 모든 시트를 검색하여 차트를 찾고, 상세한 차트 정보를 반환합니다.
    차트 관리, 대시보드 분석, 차트 인벤토리 파악에 유용합니다.

    === 워크북 접근 방법 ===
    - --file-path: 파일 경로로 워크북 열기
    - --workbook-name: 열린 워크북 이름으로 접근 (예: "Sales.xlsx")

    === 조회 범위 지정 ===
    --sheet 옵션으로 조회 범위를 제한할 수 있습니다:

    • 전체 워크북: 옵션 생략 (모든 시트의 차트 조회)
    • 특정 시트: --sheet "Dashboard" (해당 시트만 조회)
    • 여러 시트: 명령어를 여러 번 실행

    === 정보 상세도 선택 (기본값: 상세 정보 포함) ===

    ▶ 상세 정보 (기본값):
      • 차트 이름, 인덱스 번호, 위치, 크기
      • 차트 유형 (column, pie, line 등)
      • 차트 제목, 범례 설정
      • 데이터 소스 범위 (Windows만)
      • 차트 스타일 정보

    ▶ 간단 정보 (--brief 사용):
      • 차트 이름, 인덱스 번호
      • 위치 (셀 주소), 크기 (픽셀)
      • 소속 시트명

    === 활용 시나리오별 예제 ===

    # 1. 현재 워크북의 모든 차트 상세 조회 (기본)
    oa excel chart-list

    # 2. 특정 시트의 차트만 조회
    oa excel chart-list --sheet "Dashboard"

    # 3. 파일의 모든 차트 간단 조회
    oa excel chart-list --file-path "report.xlsx" --brief

    # 4. 차트 인벤토리 텍스트 형식으로 출력
    oa excel chart-list --workbook-name "Sales.xlsx" --format text

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
        # brief 옵션 처리 - 간단한 정보만 포함
        if brief:
            detailed = False

        # Engine 획득
        engine = get_engine()

        # 워크북 연결
        if file_path:
            book = engine.open_workbook(file_path, visible=visible)
        elif workbook_name:
            book = engine.get_workbook_by_name(workbook_name)
        else:
            book = engine.get_active_workbook()

        # 워크북 정보 조회
        wb_info = engine.get_workbook_info(book)

        # 차트 목록 조회 (Engine 메서드 사용)
        chart_infos = engine.list_charts(book, sheet=sheet)

        charts_info = []
        total_charts = len(chart_infos)

        # ChartInfo를 딕셔너리로 변환
        for i, chart_info in enumerate(chart_infos):
            chart_dict = {
                "index": i,
                "name": chart_info.name,
                "sheet": chart_info.sheet_name,
                "position": {"left": chart_info.left, "top": chart_info.top},
                "dimensions": {"width": chart_info.width, "height": chart_info.height},
            }

            # 상세 정보 추가
            if detailed:
                # 차트 타입 (Engine이 제공)
                chart_dict["chart_type"] = chart_info.chart_type

                # 차트 제목 (Engine이 제공)
                if chart_info.has_title and chart_info.title:
                    chart_dict["title"] = chart_info.title

                # 데이터 소스 (Engine이 제공)
                if chart_info.source_data:
                    chart_dict["data_source"] = chart_info.source_data

                # 범례 정보 (Windows COM API 직접 사용)
                if platform.system() == "Windows":
                    try:
                        ws = book.Sheets(chart_info.sheet_name)
                        chart_obj = ws.ChartObjects(chart_info.name).Chart
                        legend_info = get_chart_legend_info(chart_obj)
                        chart_dict["legend"] = legend_info
                    except:
                        chart_dict["legend"] = {"has_legend": False, "position": None}

                # 플랫폼별 추가 정보
                chart_dict["platform_support"] = {
                    "current_platform": platform.system(),
                    "full_features_available": platform.system() == "Windows",
                }

            charts_info.append(chart_dict)

        # 응답 데이터 구성
        response_data = {
            "workbook": wb_info["workbook"]["name"],
            "total_charts": total_charts,
            "charts": charts_info,
            "query_info": {
                "target_sheet": sheet if sheet else "all_sheets",
                "brief": brief,
                "detailed": detailed,
                "platform": platform.system(),
            },
        }

        if sheet:
            response_data["sheet"] = sheet
        else:
            # Engine에서 시트 개수 가져오기
            response_data["sheets_checked"] = wb_info["workbook"]["sheet_count"]

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
        # 워크북 정리 - 파일 경로로 열었고 visible=False인 경우에만 앱 종료
        if book and not visible and file_path:
            try:
                book.Application.Quit()
            except:
                pass

    return 0


if __name__ == "__main__":
    chart_list()
