"""
차트 생성 명령어 (Typer 버전)
xlwings를 활용한 Excel 차트 생성 기능
"""

import json
import platform
from pathlib import Path
from typing import Optional

import typer
import xlwings as xw

from pyhub_office_automation.version import get_version

from .utils import (
    create_error_response,
    create_success_response,
    get_or_open_workbook,
    get_range,
    get_sheet,
    normalize_path,
    parse_range,
    validate_range_string,
)

# 차트 타입 매핑 (xlwings ChartType 상수값)
CHART_TYPE_MAP = {
    "column": 51,  # xlColumnClustered
    "column_clustered": 51,
    "column_stacked": 52,  # xlColumnStacked
    "column_stacked_100": 53,  # xlColumnStacked100
    "bar": 57,  # xlBarClustered
    "bar_clustered": 57,
    "bar_stacked": 58,  # xlBarStacked
    "bar_stacked_100": 59,  # xlBarStacked100
    "line": 4,  # xlLine
    "line_markers": 65,  # xlLineMarkers
    "pie": 5,  # xlPie
    "doughnut": -4120,  # xlDoughnut
    "area": 1,  # xlArea
    "area_stacked": 76,  # xlAreaStacked
    "area_stacked_100": 77,  # xlAreaStacked100
    "scatter": -4169,  # xlXYScatter
    "scatter_lines": 74,  # xlXYScatterLines
    "scatter_smooth": 72,  # xlXYScatterSmooth
    "bubble": 15,  # xlBubble
    "combo": -4111,  # xlCombination
}


def get_chart_type_constant(chart_type: str):
    """차트 타입에 해당하는 xlwings 상수를 반환"""
    chart_type_lower = chart_type.lower()
    if chart_type_lower not in CHART_TYPE_MAP:
        raise ValueError(f"지원되지 않는 차트 타입: {chart_type}")

    # xlwings 상수를 시도하고, 실패하면 숫자값 직접 사용
    try:
        from xlwings.constants import ChartType

        # xlwings 상수명 시도
        const_map = {
            51: "xlColumnClustered",
            52: "xlColumnStacked",
            53: "xlColumnStacked100",
            57: "xlBarClustered",
            58: "xlBarStacked",
            59: "xlBarStacked100",
            4: "xlLine",
            65: "xlLineMarkers",
            5: "xlPie",
            -4120: "xlDoughnut",
            1: "xlArea",
            76: "xlAreaStacked",
            77: "xlAreaStacked100",
            -4169: "xlXYScatter",
            74: "xlXYScatterLines",
            72: "xlXYScatterSmooth",
            15: "xlBubble",
            -4111: "xlCombination",
        }

        chart_type_value = CHART_TYPE_MAP[chart_type_lower]
        const_name = const_map.get(chart_type_value)

        if const_name and hasattr(ChartType, const_name):
            return getattr(ChartType, const_name)
        else:
            # 상수 이름이 없거나 접근할 수 없으면 숫자값 직접 반환
            return chart_type_value

    except ImportError:
        # 상수를 가져올 수 없으면 숫자값 직접 반환
        return CHART_TYPE_MAP[chart_type_lower]


def chart_add(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="차트를 생성할 Excel 파일의 절대 경로"),
    use_active: bool = typer.Option(False, "--use-active", help="현재 활성 워크북 사용"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help='열린 워크북 이름으로 접근 (예: "Sales.xlsx")'),
    data_range: str = typer.Option(..., "--data-range", help='차트 데이터 범위 (예: "A1:C10" 또는 "Sheet1!A1:C10")'),
    chart_type: str = typer.Option("column", "--chart-type", help="차트 유형 (기본값: column)"),
    title: Optional[str] = typer.Option(None, "--title", help="차트 제목"),
    position: str = typer.Option("E1", "--position", help="차트 생성 위치 (셀 주소, 기본값: E1)"),
    width: int = typer.Option(400, "--width", help="차트 너비 (픽셀, 기본값: 400)"),
    height: int = typer.Option(300, "--height", help="차트 높이 (픽셀, 기본값: 300)"),
    sheet: Optional[str] = typer.Option(None, "--sheet", help="차트를 생성할 시트 이름 (지정하지 않으면 데이터 범위의 시트)"),
    style: Optional[int] = typer.Option(None, "--style", help="차트 스타일 번호 (1-48)"),
    legend_position: Optional[str] = typer.Option(None, "--legend-position", help="범례 위치 (top/bottom/left/right/none)"),
    show_data_labels: bool = typer.Option(False, "--show-data-labels", help="데이터 레이블 표시"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택 (json/text)"),
    visible: bool = typer.Option(False, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부 (기본값: False)"),
    save: bool = typer.Option(True, "--save", help="생성 후 파일 저장 여부 (기본값: True)"),
):
    """
    지정된 데이터 범위에서 Excel 차트를 생성합니다.

    다양한 차트 유형을 지원하며, 위치와 크기를 정밀하게 제어할 수 있습니다.
    Windows와 macOS 모두에서 동작하지만, 일부 고급 기능은 Windows에서만 지원됩니다.

    === 워크북 접근 방법 ===
    - --file-path: 파일 경로로 워크북 열기
    - --use-active: 현재 활성 워크북 사용
    - --workbook-name: 열린 워크북 이름으로 접근 (예: "Sales.xlsx")

    === 데이터 범위 지정 방법 ===
    --data-range 옵션으로 차트 데이터를 지정합니다:

    • 현재 시트 범위: "A1:C10"
    • 특정 시트 범위: "Sheet1!A1:C10"
    • 공백 포함 시트명: "'데이터 시트'!A1:C10"
    • 헤더 포함 권장: 첫 행은 열 제목, 나머지는 데이터

    === 차트 위치 지정 방법 ===
    --position과 --sheet 옵션으로 차트 위치를 지정합니다:

    • 기본 위치: E1 (지정하지 않은 경우)
    • 셀 주소 지정: --position "H5" (H열 5행)
    • 다른 시트에 생성: --sheet "Dashboard" --position "B2"
    • 새 시트 자동 생성: 지정한 시트가 없으면 자동으로 생성

    === 지원되는 차트 유형과 적합한 데이터 구조 ===

    ▶ 원형/도넛 차트 (pie, doughnut):
      • 데이터 구조: [레이블, 값] - 2열 필요
      • 예: A열=제품명, B열=판매량

    ▶ 막대/선 차트 (column, bar, line):
      • 데이터 구조: [카테고리, 시리즈1, 시리즈2, ...]
      • 예: A열=월, B열=매출, C열=비용

    ▶ 산점도 (scatter):
      • 데이터 구조: [X값, Y값] 또는 [X값, Y값, 크기]
      • 예: A열=광고비, B열=매출

    ▶ 버블 차트 (bubble):
      • 데이터 구조: [X값, Y값, 크기값]
      • 예: A열=가격, B열=품질점수, C열=판매량

    === 차트 스타일링 옵션 ===
    --style: 차트 스타일 번호 (1-48)
    --legend-position: 범례 위치 (top/bottom/left/right/none)
    --show-data-labels: 데이터 레이블 표시
    --title: 차트 제목 설정

    === 실제 사용 시나리오 예제 ===

    # 1. 기본 매출 차트 생성
    oa excel chart-add --use-active --data-range "A1:C10" --chart-type "column" --title "매출 현황"

    # 2. 특정 시트 데이터로 원형 차트 생성
    oa excel chart-add --file-path "sales.xlsx" --data-range "Sheet1!A1:D20" --chart-type "pie" --position "F5"

    # 3. 대시보드용 차트를 별도 시트에 생성
    oa excel chart-add --use-active --data-range "Data!A1:E15" --sheet "Dashboard" --position "B2" --chart-type "line"

    # 4. 스타일링이 적용된 차트 생성
    oa excel chart-add --workbook-name "Report.xlsx" --data-range "A1:C15" --chart-type "column" \\
        --title "월별 실적" --style 10 --legend-position "bottom" --show-data-labels --position "H1"

    # 5. 산점도로 상관관계 분석
    oa excel chart-add --use-active --data-range "Analysis!A1:C20" --chart-type "scatter" \\
        --title "광고비 vs 매출 상관관계" --position "F10"
    """
    book = None

    try:
        # chart_type 검증
        valid_chart_types = [
            "column",
            "column_clustered",
            "column_stacked",
            "column_stacked_100",
            "bar",
            "bar_clustered",
            "bar_stacked",
            "bar_stacked_100",
            "line",
            "line_markers",
            "pie",
            "doughnut",
            "area",
            "area_stacked",
            "area_stacked_100",
            "scatter",
            "scatter_lines",
            "scatter_smooth",
            "bubble",
            "combo",
        ]
        if chart_type not in valid_chart_types:
            raise ValueError(f"지원되지 않는 차트 타입: {chart_type}. 사용 가능한 타입: {', '.join(valid_chart_types)}")

        # legend_position 검증
        if legend_position and legend_position not in ["top", "bottom", "left", "right", "none"]:
            raise ValueError(f"잘못된 범례 위치: {legend_position}. 사용 가능한 위치: top, bottom, left, right, none")

        # output_format 검증
        if output_format not in ["json", "text"]:
            raise ValueError(f"잘못된 출력 형식: {output_format}. 사용 가능한 형식: json, text")

        # 데이터 범위 파싱 및 검증
        data_sheet_name, data_range_part = parse_range(data_range)
        if not validate_range_string(data_range_part):
            raise ValueError(f"잘못된 데이터 범위 형식입니다: {data_range}")

        # 워크북 연결
        book = get_or_open_workbook(file_path=file_path, workbook_name=workbook_name, use_active=use_active, visible=visible)

        # 데이터 시트 가져오기
        data_sheet = get_sheet(book, data_sheet_name)

        # 데이터 범위 가져오기 및 검증
        data_chart_range = get_range(data_sheet, data_range_part)
        data_values = data_chart_range.value

        if not data_values or (isinstance(data_values, list) and len(data_values) == 0):
            raise ValueError("데이터 범위에 차트 생성을 위한 데이터가 없습니다")

        # 차트 생성 대상 시트 결정
        if sheet:
            try:
                target_sheet = get_sheet(book, sheet)
            except ValueError:
                # 지정한 시트가 없으면 새로 생성
                target_sheet = book.sheets.add(name=sheet)
        else:
            # 시트가 지정되지 않으면 데이터가 있는 시트 사용
            target_sheet = data_sheet

        # 차트 생성 위치 결정
        try:
            position_range = target_sheet.range(position)
            left = position_range.left
            top = position_range.top
        except Exception:
            # 잘못된 위치가 지정된 경우 기본 위치 사용
            left = 300
            top = 50

        # 차트 타입 상수 가져오기
        try:
            chart_type_const = get_chart_type_constant(chart_type)
        except Exception as e:
            raise ValueError(f"차트 타입 처리 오류: {str(e)}")

        # 차트 생성
        try:
            # xlwings 방식: 먼저 차트 객체를 생성하고 나중에 데이터 설정
            chart = target_sheet.charts.add(left=left, top=top, width=width, height=height)

            # 차트에 데이터 범위 설정
            chart.set_source_data(data_chart_range)

            # 차트 타입 설정
            try:
                if platform.system() == "Windows":
                    # Windows에서는 API를 통해 직접 설정
                    chart.api.ChartType = chart_type_const
                else:
                    # macOS에서는 chart_type 속성 사용 (제한적)
                    chart.chart_type = chart_type_const
            except:
                # 차트 타입 설정 실패 시 기본값 유지
                pass

            chart_name = chart.name

        except Exception as e:
            raise RuntimeError(f"차트 생성 실패: {str(e)}")

        # 차트 제목 설정
        if title:
            try:
                chart.api.HasTitle = True
                chart.api.ChartTitle.Text = title
            except:
                # 제목 설정 실패해도 계속 진행
                pass

        # 차트 스타일 설정 (Windows에서만 가능)
        if style and platform.system() == "Windows":
            try:
                chart.api.ChartStyle = style
            except:
                pass

        # 범례 위치 설정
        if legend_position:
            try:
                if legend_position == "none":
                    chart.api.HasLegend = False
                else:
                    chart.api.HasLegend = True
                    if platform.system() == "Windows":
                        from xlwings.constants import LegendPosition

                        legend_map = {
                            "top": LegendPosition.xlLegendPositionTop,
                            "bottom": LegendPosition.xlLegendPositionBottom,
                            "left": LegendPosition.xlLegendPositionLeft,
                            "right": LegendPosition.xlLegendPositionRight,
                        }
                        if legend_position in legend_map:
                            chart.api.Legend.Position = legend_map[legend_position]
            except:
                pass

        # 데이터 레이블 표시
        if show_data_labels and platform.system() == "Windows":
            try:
                chart.api.FullSeriesCollection(1).HasDataLabels = True
            except:
                pass

        # 파일 저장
        if save and file_path:
            book.save()

        # 성공 응답 생성
        response_data = {
            "chart_name": chart_name,
            "chart_type": chart_type,
            "data_range": data_range,
            "position": position,
            "dimensions": {"width": width, "height": height},
            "sheet": target_sheet.name,
            "workbook": book.name,
        }

        if title:
            response_data["title"] = title

        response = create_success_response(
            data=response_data, command="chart-add", message=f"차트 '{chart_name}'이 성공적으로 생성되었습니다"
        )

        print(json.dumps(response, ensure_ascii=False, indent=2))

    except Exception as e:
        error_response = create_error_response(e, "chart-add")
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
    chart_add()
