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
    ChartType,
    ExpandMode,
    LegendPosition,
    OutputFormat,
    check_range_overlap,
    coords_to_excel_address,
    create_error_response,
    create_success_response,
    find_available_position,
    get_all_chart_ranges,
    get_all_pivot_ranges,
    get_chart_com_object,
    get_or_open_workbook,
    get_range,
    get_sheet,
    normalize_path,
    parse_range,
    validate_auto_position_requirements,
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


def get_chart_type_constant(chart_type: ChartType):
    """차트 타입에 해당하는 xlwings 상수를 반환"""
    chart_type_value = chart_type.value
    if chart_type_value not in CHART_TYPE_MAP:
        raise ValueError(f"지원되지 않는 차트 타입: {chart_type}")

    # xlwings 상수를 시도하고, 실패하면 숫자값 직접 사용
    try:
        from xlwings.constants import ChartType as XlChartType

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

        chart_type_code = CHART_TYPE_MAP[chart_type_value]
        const_name = const_map.get(chart_type_code)

        if const_name and hasattr(XlChartType, const_name):
            return getattr(XlChartType, const_name)
        else:
            # 상수 이름이 없거나 접근할 수 없으면 숫자값 직접 반환
            return chart_type_code

    except ImportError:
        # 상수를 가져올 수 없으면 숫자값 직접 반환
        return CHART_TYPE_MAP[chart_type_value]


def chart_add(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="차트를 생성할 Excel 파일의 절대 경로"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help='열린 워크북 이름으로 접근 (예: "Sales.xlsx")'),
    data_range: str = typer.Option(..., "--data-range", help='차트 데이터 범위 (예: "A1:C10" 또는 "Sheet1!A1:C10")'),
    expand: Optional[ExpandMode] = typer.Option(None, "--expand", help="데이터 범위 확장 모드 (table, down, right)"),
    chart_type: ChartType = typer.Option(ChartType.COLUMN, "--chart-type", help="차트 유형 (기본값: column)"),
    title: Optional[str] = typer.Option(None, "--title", help="차트 제목"),
    position: str = typer.Option("E1", "--position", help="차트 생성 위치 (셀 주소, 기본값: E1)"),
    width: int = typer.Option(400, "--width", help="차트 너비 (픽셀, 기본값: 400)"),
    height: int = typer.Option(300, "--height", help="차트 높이 (픽셀, 기본값: 300)"),
    sheet: Optional[str] = typer.Option(None, "--sheet", help="차트를 생성할 시트 이름 (지정하지 않으면 데이터 범위의 시트)"),
    auto_position: bool = typer.Option(False, "--auto-position", help="자동으로 빈 공간을 찾아 배치"),
    check_overlap: bool = typer.Option(False, "--check-overlap", help="지정된 위치의 겹침 검사 후 경고 표시"),
    spacing: int = typer.Option(50, "--spacing", help="자동 배치 시 기존 객체와의 최소 간격 (픽셀 단위, 기본값: 50)"),
    preferred_position: str = typer.Option(
        "right", "--preferred-position", help="자동 배치 시 선호 방향 (right/bottom, 기본값: right)"
    ),
    style: Optional[int] = typer.Option(None, "--style", help="차트 스타일 번호 (1-48)"),
    legend_position: Optional[LegendPosition] = typer.Option(
        None, "--legend-position", help="범례 위치 (top/bottom/left/right/none)"
    ),
    show_data_labels: bool = typer.Option(False, "--show-data-labels", help="데이터 레이블 표시"),
    output_format: OutputFormat = typer.Option(OutputFormat.JSON, "--format", help="출력 형식 선택 (json/text)"),
    visible: bool = typer.Option(False, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부 (기본값: False)"),
    save: bool = typer.Option(True, "--save", help="생성 후 파일 저장 여부 (기본값: True)"),
):
    """
    지정된 데이터 범위에서 Excel 차트를 생성합니다.

    다양한 차트 유형을 지원하며, 위치와 크기를 정밀하게 제어할 수 있습니다.
    Windows와 macOS 모두에서 동작하지만, 일부 고급 기능은 Windows에서만 지원됩니다.

    \b
    워크북 접근 방법:
      • 옵션 없음: 활성 워크북 자동 사용 (기본값)
      • --file-path: 파일 경로로 워크북 열기
      • --workbook-name: 열린 워크북 이름으로 접근

    \b
    데이터 범위 지정 방법:
      --data-range 옵션으로 차트 데이터를 지정합니다:
      • 현재 시트 범위: "A1:C10"
      • 특정 시트 범위: "Sheet1!A1:C10"
      • 공백 포함 시트명: "'데이터 시트'!A1:C10"
      • 헤더 포함 권장: 첫 행은 열 제목, 나머지는 데이터

    \b
    데이터 범위 확장 모드:
      • --expand table: 연결된 데이터 테이블 전체로 확장
      • --expand down: 아래쪽으로 데이터가 있는 곳까지 확장
      • --expand right: 오른쪽으로 데이터가 있는 곳까지 확장

    \b
    차트 위치 지정 방법:
      --position과 --sheet 옵션으로 차트 위치를 지정합니다:
      • 기본 위치: E1 (지정하지 않은 경우)
      • 셀 주소 지정: --position "H5" (H열 5행)
      • 다른 시트에 생성: --sheet "Dashboard" --position "B2"
      • 새 시트 자동 생성: 지정한 시트가 없으면 자동으로 생성

    \b
    자동 배치 기능:
      • --auto-position: 기존 피벗테이블과 차트를 피해 자동으로 빈 공간 찾기
      • --check-overlap: 지정된 위치가 기존 객체와 겹치는지 검사
      • --spacing: 자동 배치 시 최소 간격 설정 (픽셀 단위, 기본값: 50)
      • --preferred-position: 배치 방향 선호도 (right/bottom)

    \b
    지원되는 차트 유형과 적합한 데이터 구조:
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

    \b
    차트 스타일링 옵션:
      • --style: 차트 스타일 번호 (1-48)
      • --legend-position: 범례 위치 (top/bottom/left/right/none)
      • --show-data-labels: 데이터 레이블 표시
      • --title: 차트 제목 설정

    \b
    사용 예제:
      # 기본 매출 차트 생성
      oa excel chart-add --data-range "A1:C10" --chart-type "column" --title "매출 현황"

      # 특정 시트 데이터로 원형 차트 생성
      oa excel chart-add --file-path "sales.xlsx" --data-range "Sheet1!A1:D20" --chart-type "pie" --position "F5"

      # 자동 배치로 차트 생성 (첫 번째 차트 후 사용 권장)
      oa excel chart-add --data-range "A1:C10" --chart-type "column" --auto-position --title "매출 현황"

      # 자동 배치 + 사용자 설정
      oa excel chart-add --data-range "A1:C10" --auto-position --spacing 80 --preferred-position "bottom" --chart-type "line"

      # 겹침 검사
      oa excel chart-add --data-range "A1:C10" --position "F5" --check-overlap --chart-type "pie"

      # 데이터 범위 자동 확장으로 차트 생성
      oa excel chart-add --data-range "A1" --expand table --chart-type "column" --title "전체 데이터 차트"

      # 대시보드용 차트를 별도 시트에 생성
      oa excel chart-add --data-range "Data!A1:E15" --sheet "Dashboard" --position "B2" --chart-type "line"

      # 스타일링이 적용된 차트 자동 배치
      oa excel chart-add --workbook-name "Report.xlsx" --data-range "A1" --expand table --chart-type "column" \\
          --title "월별 실적" --style 10 --legend-position "bottom" --show-data-labels --auto-position
    """
    book = None

    try:
        # Enum 타입이므로 별도 검증 불필요 (chart_type은 ChartType Enum이므로 자동 검증됨)
        # legend_position도 Enum 타입이므로 별도 검증 불필요
        # output_format도 Enum 타입이므로 별도 검증 불필요

        # 자동 배치와 수동 배치 옵션 충돌 검사
        if auto_position and position != "E1":
            raise ValueError("--auto-position 옵션 사용 시 --position을 지정할 수 없습니다. 자동으로 위치가 결정됩니다.")

        # preferred_position 검증
        if preferred_position not in ["right", "bottom"]:
            raise ValueError("--preferred-position은 'right' 또는 'bottom'만 지원됩니다.")

        # spacing 검증 (픽셀 단위)
        if spacing < 10 or spacing > 200:
            raise ValueError("--spacing은 10~200 픽셀 사이의 값이어야 합니다.")

        # 데이터 범위 파싱 및 검증
        data_sheet_name, data_range_part = parse_range(data_range)
        if not validate_range_string(data_range_part):
            raise ValueError(f"잘못된 데이터 범위 형식입니다: {data_range}")

        # 워크북 연결
        book = get_or_open_workbook(file_path=file_path, workbook_name=workbook_name, visible=visible)

        # 데이터 시트 가져오기
        data_sheet = get_sheet(book, data_sheet_name)

        # 데이터 범위 가져오기 및 검증 (expand 옵션 적용)
        data_chart_range = get_range(data_sheet, data_range_part, expand_mode=expand)
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

        # 자동 배치 또는 수동 배치 처리
        overlap_warning = None
        auto_position_info = None

        if auto_position:
            # 자동 배치 기능 사용 가능 여부 확인
            can_auto_position, auto_error = validate_auto_position_requirements(target_sheet)
            if not can_auto_position:
                # 차트는 피벗테이블과 달리 macOS에서도 생성 가능하므로 경고만 표시
                print(f"⚠️ 자동 배치 제한: {auto_error}")
                print("💡 macOS에서는 기본 위치(E1)를 사용합니다.")
                position_range = target_sheet.range("E1")
                left = position_range.left
                top = position_range.top
            else:
                # 차트 크기를 고려한 예상 범위 계산
                # 차트는 픽셀 단위이므로 셀 단위로 변환
                chart_cols = max(1, int(width / 64))  # Excel 기본 열 너비 64픽셀
                chart_rows = max(1, int(height / 15))  # Excel 기본 행 높이 15픽셀

                # 자동으로 빈 위치 찾기
                try:
                    auto_dest_range = find_available_position(
                        target_sheet,
                        min_spacing=max(1, int(spacing / 64)),  # 픽셀을 열 단위로 변환
                        preferred_position=preferred_position,
                        estimate_size=(chart_cols, chart_rows),
                    )
                    position_range = target_sheet.range(auto_dest_range)
                    left = position_range.left
                    top = position_range.top

                    auto_position_info = {
                        "original_request": "auto",
                        "found_position": auto_dest_range,
                        "estimated_size": {"cols": chart_cols, "rows": chart_rows},
                        "spacing_used": spacing,
                        "preferred_direction": preferred_position,
                    }
                except RuntimeError as e:
                    raise RuntimeError(f"자동 배치 실패: {str(e)}")

        else:
            # 수동 배치: 기존 로직 사용
            try:
                position_range = target_sheet.range(position)
                left = position_range.left
                top = position_range.top
            except Exception:
                # 잘못된 위치가 지정된 경우 기본 위치 사용
                left = 300
                top = 50

            # 겹침 검사 옵션 처리
            if check_overlap:
                # 차트 크기를 고려한 예상 범위 계산
                chart_cols = max(1, int(width / 64))
                chart_rows = max(1, int(height / 15))

                # 시작 위치에서 차트 크기만큼의 범위 계산
                start_col = max(1, int(left / 64) + 1)
                start_row = max(1, int(top / 15) + 1)
                end_col = start_col + chart_cols - 1
                end_row = start_row + chart_rows - 1

                estimated_range = (
                    f"{coords_to_excel_address(start_row, start_col)}:{coords_to_excel_address(end_row, end_col)}"
                )

                # 기존 피벗 테이블 범위 확인
                existing_pivots = get_all_pivot_ranges(target_sheet)
                overlapping_pivots = []

                for pivot_range in existing_pivots:
                    if check_range_overlap(estimated_range, pivot_range):
                        overlapping_pivots.append(pivot_range)

                # 기존 차트 범위 확인
                chart_info = get_all_chart_ranges(target_sheet)
                overlapping_charts = []

                for chart_range, _, _ in chart_info:
                    if check_range_overlap(estimated_range, chart_range):
                        overlapping_charts.append(chart_range)

                if overlapping_pivots or overlapping_charts:
                    overlap_warning = {
                        "estimated_range": estimated_range,
                        "overlapping_pivots": overlapping_pivots,
                        "overlapping_charts": overlapping_charts,
                        "recommendation": "다른 위치를 선택하거나 --auto-position 옵션을 사용하세요.",
                    }

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
                # 실제 Chart COM 객체 가져오기
                chart_com = get_chart_com_object(chart)

                if platform.system() == "Windows":
                    # Windows에서는 API를 통해 직접 설정
                    chart_com.ChartType = chart_type_const
                else:
                    # macOS에서는 chart_type 속성 사용 (제한적)
                    chart.chart_type = chart_type_const
            except:
                # 차트 타입 설정 실패 시 기본값 유지
                pass

            chart_name = chart.name

        except Exception as e:
            # COM 에러인 경우 더 자세한 처리를 위해 그대로 전달
            if "com_error" in str(type(e).__name__).lower():
                raise
            else:
                raise RuntimeError(f"차트 생성 실패: {str(e)}")

        # 차트 제목 설정
        if title:
            try:
                # 실제 Chart COM 객체 가져오기
                chart_com = get_chart_com_object(chart)
                chart_com.HasTitle = True
                chart_com.ChartTitle.Text = title
            except:
                # 제목 설정 실패해도 계속 진행
                pass

        # 차트 스타일 설정 (Windows에서만 가능)
        if style and platform.system() == "Windows":
            try:
                # 실제 Chart COM 객체 가져오기
                chart_com = get_chart_com_object(chart)
                chart_com.ChartStyle = style
            except:
                pass

        # 범례 위치 설정
        if legend_position:
            try:
                # 실제 Chart COM 객체 가져오기
                chart_com = get_chart_com_object(chart)

                # legend_position을 문자열로 정규화
                if hasattr(legend_position, 'value'):
                    position_str = legend_position.value
                else:
                    position_str = str(legend_position).lower()

                if position_str == "none":
                    chart_com.HasLegend = False
                else:
                    chart_com.HasLegend = True
                    if platform.system() == "Windows":
                        legend_map = {
                            "top": -4160,  # xlLegendPositionTop
                            "bottom": -4107,  # xlLegendPositionBottom
                            "left": -4131,  # xlLegendPositionLeft
                            "right": -4152,  # xlLegendPositionRight
                        }
                        if position_str in legend_map:
                            chart_com.Legend.Position = legend_map[position_str]
            except:
                pass

        # 데이터 레이블 표시
        if show_data_labels and platform.system() == "Windows":
            try:
                # 실제 Chart COM 객체 가져오기
                chart_com = get_chart_com_object(chart)
                chart_com.FullSeriesCollection(1).HasDataLabels = True
            except:
                pass

        # 파일 저장
        if save and file_path:
            book.save()

        # 성공 응답 생성
        response_data = {
            "chart_name": chart_name,
            "chart_type": str(chart_type),
            "data_range": data_range,
            "position": position_range.address if auto_position else position,
            "dimensions": {"width": width, "height": height},
            "sheet": target_sheet.name,
            "workbook": book.name,
        }

        if title:
            response_data["title"] = title

        # 자동 배치 정보 추가
        if auto_position_info:
            response_data["auto_position"] = auto_position_info

        # 겹침 경고 추가
        if overlap_warning:
            response_data["overlap_warning"] = overlap_warning

        response = create_success_response(
            data=response_data, command="chart-add", message=f"차트 '{chart_name}'이 성공적으로 생성되었습니다"
        )

        # JSON 출력 시 자동 배치/겹침 정보도 포함하여 출력
        if output_format == OutputFormat.JSON:
            print(json.dumps(response, ensure_ascii=False, indent=2))
        else:
            # text 출력 형식에서도 자동 배치 정보 표시
            print(f"✅ 차트 생성 성공")
            print(f"📊 차트 이름: {chart_name}")
            print(f"📄 파일: {book.name}")
            print(f"📈 데이터 범위: {data_range}")
            print(f"📍 생성 위치: {target_sheet.name}!{response_data['position']}")
            print(f"📏 크기: {width}×{height} 픽셀")

            # 자동 배치 정보 표시
            if auto_position_info:
                print(
                    f"🎯 자동 배치: {auto_position_info['found_position']} (방향: {auto_position_info['preferred_direction']}, 간격: {auto_position_info['spacing_used']}px)"
                )
                print(
                    f"📐 예상 크기: {auto_position_info['estimated_size']['cols']}열 × {auto_position_info['estimated_size']['rows']}행"
                )

            # 겹침 경고 표시
            if overlap_warning:
                print("⚠️  겹침 경고!")
                print(f"   예상 범위: {overlap_warning['estimated_range']}")
                if overlap_warning["overlapping_pivots"]:
                    print(f"   겹치는 피벗테이블: {', '.join(overlap_warning['overlapping_pivots'])}")
                if overlap_warning["overlapping_charts"]:
                    print(f"   겹치는 차트: {len(overlap_warning['overlapping_charts'])}개")
                print(f"   💡 {overlap_warning['recommendation']}")

            if title:
                print(f"📝 제목: {title}")

    except Exception as e:
        error_response = create_error_response(e, "chart-add")
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
    typer.run(chart_add)
