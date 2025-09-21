"""
차트 설정 명령어
기존 차트의 스타일과 속성을 설정하는 기능
"""

import json
import platform
from pathlib import Path
from typing import Optional

import typer
import xlwings as xw

from pyhub_office_automation.version import get_version

from .utils import create_error_response, create_success_response, get_or_open_workbook, get_sheet, normalize_path


def find_chart_by_name_or_index(sheet, chart_name=None, chart_index=None):
    """차트 이름이나 인덱스로 차트 객체 찾기"""
    if chart_name:
        for chart in sheet.charts:
            if chart.name == chart_name:
                return chart
        raise ValueError(f"차트 '{chart_name}'을 찾을 수 없습니다")

    elif chart_index is not None:
        try:
            if 0 <= chart_index < len(sheet.charts):
                return sheet.charts[chart_index]
            else:
                raise IndexError(f"차트 인덱스 {chart_index}는 범위를 벗어났습니다 (0-{len(sheet.charts)-1})")
        except IndexError as e:
            raise ValueError(str(e))

    else:
        raise ValueError("차트 이름(--chart-name) 또는 인덱스(--chart-index) 중 하나를 지정해야 합니다")


def set_chart_style(chart, style_number):
    """차트 스타일 설정 (Windows 전용)"""
    if platform.system() != "Windows":
        return False

    try:
        if 1 <= style_number <= 48:
            chart.api.ChartStyle = style_number
            return True
        else:
            raise ValueError("차트 스타일은 1-48 범위여야 합니다")
    except Exception:
        return False


def set_legend_position(chart, position):
    """범례 위치 설정"""
    try:
        if position == "none":
            chart.api.HasLegend = False
            return True

        chart.api.HasLegend = True

        if platform.system() == "Windows":
            # LegendPosition 상수값 직접 사용
            position_map = {
                "top": -4160,  # xlLegendPositionTop
                "bottom": -4107,  # xlLegendPositionBottom
                "left": -4152,  # xlLegendPositionLeft
                "right": -4161,  # xlLegendPositionRight
            }

            if position in position_map:
                try:
                    # xlwings 상수 시도
                    from xlwings.constants import LegendPosition

                    const_map = {
                        -4160: "xlLegendPositionTop",
                        -4107: "xlLegendPositionBottom",
                        -4152: "xlLegendPositionLeft",
                        -4161: "xlLegendPositionRight",
                    }
                    position_value = position_map[position]
                    const_name = const_map.get(position_value)

                    if const_name and hasattr(LegendPosition, const_name):
                        chart.api.Legend.Position = getattr(LegendPosition, const_name)
                    else:
                        # 상수를 찾을 수 없으면 숫자값 직접 사용
                        chart.api.Legend.Position = position_value
                except ImportError:
                    # LegendPosition을 가져올 수 없으면 숫자값 직접 사용
                    chart.api.Legend.Position = position_map[position]
                return True

        return False
    except Exception:
        return False


def set_axis_titles(chart, x_title=None, y_title=None):
    """축 제목 설정 (Windows에서 더 안정적)"""
    results = {"x_axis": False, "y_axis": False}

    try:
        if x_title:
            chart.api.Axes(1).HasTitle = True  # 1 = X축
            chart.api.Axes(1).AxisTitle.Text = x_title
            results["x_axis"] = True
    except Exception:
        pass

    try:
        if y_title:
            chart.api.Axes(2).HasTitle = True  # 2 = Y축
            chart.api.Axes(2).AxisTitle.Text = y_title
            results["y_axis"] = True
    except Exception:
        pass

    return results


def set_data_labels(chart, show_labels, label_position=None):
    """데이터 레이블 설정"""
    try:
        if platform.system() == "Windows":
            series_collection = chart.api.FullSeriesCollection()
            for i in range(1, series_collection.Count + 1):
                series = series_collection(i)
                series.HasDataLabels = show_labels

                if show_labels and label_position:
                    # 레이블 위치 설정 (Windows 전용) - 상수값 직접 사용
                    position_map = {
                        "center": -4108,  # xlLabelPositionCenter
                        "above": -4117,  # xlLabelPositionAbove
                        "below": -4107,  # xlLabelPositionBelow
                        "left": -4131,  # xlLabelPositionLeft
                        "right": -4152,  # xlLabelPositionRight
                        "outside": -4114,  # xlLabelPositionOutsideEnd
                        "inside": -4112,  # xlLabelPositionInsideEnd
                    }

                    if label_position in position_map:
                        try:
                            # xlwings 상수 시도
                            from xlwings.constants import DataLabelPosition

                            const_map = {
                                -4108: "xlLabelPositionCenter",
                                -4117: "xlLabelPositionAbove",
                                -4107: "xlLabelPositionBelow",
                                -4131: "xlLabelPositionLeft",
                                -4152: "xlLabelPositionRight",
                                -4114: "xlLabelPositionOutsideEnd",
                                -4112: "xlLabelPositionInsideEnd",
                            }
                            position_value = position_map[label_position]
                            const_name = const_map.get(position_value)

                            if const_name and hasattr(DataLabelPosition, const_name):
                                series.DataLabels().Position = getattr(DataLabelPosition, const_name)
                            else:
                                # 상수를 찾을 수 없으면 숫자값 직접 사용
                                series.DataLabels().Position = position_value
                        except ImportError:
                            # DataLabelPosition을 가져올 수 없으면 숫자값 직접 사용
                            series.DataLabels().Position = position_map[label_position]
                        except:
                            pass

            return True
        else:
            # macOS에서는 기본적인 설정만 가능
            return False
    except Exception:
        return False


def set_chart_colors(chart, color_scheme):
    """차트 색상 테마 설정 (Windows에서 더 많은 옵션)"""
    try:
        if platform.system() == "Windows":
            # 색상 스키마 적용
            color_schemes = {"colorful": 2, "monochromatic": 3, "office": 1, "grayscale": 4}

            if color_scheme in color_schemes:
                chart.api.ChartColorIndex = color_schemes[color_scheme]
                return True

        return False
    except Exception:
        return False


def chart_configure(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="차트가 있는 Excel 파일의 절대 경로"),
    use_active: bool = typer.Option(False, "--use-active", help="현재 활성 워크북 사용"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help='열린 워크북 이름으로 접근 (예: "Sales.xlsx")'),
    sheet: Optional[str] = typer.Option(None, "--sheet", help="차트가 있는 시트 이름 (지정하지 않으면 활성 시트)"),
    chart_name: Optional[str] = typer.Option(None, "--chart-name", help="설정할 차트의 이름"),
    chart_index: Optional[int] = typer.Option(None, "--chart-index", help="설정할 차트의 인덱스 (0부터 시작)"),
    title: Optional[str] = typer.Option(None, "--title", help="차트 제목 설정"),
    style: Optional[int] = typer.Option(None, "--style", help="차트 스타일 번호 (1-48, Windows 전용)"),
    legend_position: Optional[str] = typer.Option(None, "--legend-position", help="범례 위치 (top/bottom/left/right/none)"),
    x_axis_title: Optional[str] = typer.Option(None, "--x-axis-title", help="X축 제목"),
    y_axis_title: Optional[str] = typer.Option(None, "--y-axis-title", help="Y축 제목"),
    show_data_labels: bool = typer.Option(False, "--show-data-labels", help="데이터 레이블 표시"),
    hide_data_labels: bool = typer.Option(False, "--hide-data-labels", help="데이터 레이블 숨기기"),
    data_label_position: Optional[str] = typer.Option(
        None, "--data-label-position", help="데이터 레이블 위치 (center/above/below/left/right/outside/inside, Windows 전용)"
    ),
    color_scheme: Optional[str] = typer.Option(
        None, "--color-scheme", help="색상 테마 (colorful/monochromatic/office/grayscale, Windows 전용)"
    ),
    transparent_bg: bool = typer.Option(False, "--transparent-bg", help="차트 배경을 투명하게 설정"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택 (json/text)"),
    visible: bool = typer.Option(False, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부 (기본값: False)"),
    save: bool = typer.Option(True, "--save", help="설정 후 파일 저장 여부 (기본값: True)"),
):
    """
    기존 차트의 스타일과 속성을 설정합니다.

    생성된 차트의 외관과 속성을 세밀하게 조정할 수 있습니다. 차트 제목, 스타일, 범례, 축 제목,
    데이터 레이블 등을 설정하여 전문적인 차트를 완성할 수 있습니다.

    === 워크북 접근 방법 ===
    - --file-path: 파일 경로로 워크북 열기
    - --use-active: 현재 활성 워크북 사용 (권장)
    - --workbook-name: 열린 워크북 이름으로 접근 (예: "Sales.xlsx")

    === 차트 선택 방법 ===
    차트를 선택하는 두 가지 방법이 있습니다:

    ▶ 차트 이름으로 선택:
      • --chart-name "Chart1"
      • chart-list 명령으로 차트 이름 확인 가능

    ▶ 인덱스 번호로 선택:
      • --chart-index 0 (첫 번째 차트)
      • 시트의 차트 순서대로 0, 1, 2...

    === 설정 가능한 속성 ===

    ▶ 제목 및 레이블:
      • --title "새 제목": 차트 제목 설정
      • --x-axis-title "X축 제목": X축 제목
      • --y-axis-title "Y축 제목": Y축 제목

    ▶ 스타일 및 외관:
      • --style 1-48: 차트 스타일 번호 (Windows 전용)
      • --color-scheme: colorful/monochromatic/office/grayscale
      • --transparent-bg: 차트 배경 투명화

    ▶ 범례 설정:
      • --legend-position: top/bottom/left/right/none
      • 범례 위치를 조정하여 차트 가독성 향상

    ▶ 데이터 레이블:
      • --show-data-labels: 데이터 레이블 표시
      • --hide-data-labels: 데이터 레이블 숨김
      • --data-label-position: center/above/below/left/right/outside/inside (Windows)

    === 차트 유형별 권장 설정 ===

    ▶ 막대/선 차트:
      • 축 제목 추가로 데이터 의미 명확화
      • 범례를 하단에 배치하여 공간 효율성
      • 데이터 레이블은 선택적 사용

    ▶ 원형/도넛 차트:
      • 데이터 레이블 표시 권장 (백분율 또는 값)
      • 범례를 우측에 배치
      • 색상 구분이 중요하므로 colorful 테마 사용

    ▶ 산점도:
      • X, Y축 제목 필수 (상관관계 표현)
      • 범례 위치는 데이터 분포에 따라 조정
      • 투명 배경으로 데이터 포인트 강조

    === 실제 활용 시나리오 예제 ===

    # 1. 기본 차트 스타일링
    oa excel chart-configure --use-active --chart-index 0 --title "2024년 매출 현황" --legend-position "bottom"

    # 2. 축 제목과 데이터 레이블 추가
    oa excel chart-configure --use-active --chart-name "SalesChart" \\
        --x-axis-title "월" --y-axis-title "매출액(만원)" --show-data-labels

    # 3. 프레젠테이션용 고급 스타일링 (Windows)
    oa excel chart-configure --file-path "report.xlsx" --chart-index 0 \\
        --style 15 --color-scheme "office" --transparent-bg --data-label-position "outside"

    # 4. 원형 차트 최적화
    oa excel chart-configure --workbook-name "Dashboard.xlsx" --chart-name "MarketShare" \\
        --title "시장 점유율" --legend-position "right" --show-data-labels

    # 5. 산점도 상관관계 차트 설정
    oa excel chart-configure --use-active --chart-index 2 \\
        --title "광고비 vs 매출 상관관계" --x-axis-title "광고비(만원)" --y-axis-title "매출(억원)"

    === 플랫폼별 기능 차이 ===
    • Windows: 모든 설정 옵션 지원 (스타일, 색상 테마, 레이블 위치 등)
    • macOS: 기본 설정만 지원 (제목, 범례 위치, 레이블 표시/숨김)

    === 팁 ===
    • chart-list --detailed로 현재 설정 확인 후 수정
    • 한 번에 여러 속성을 동시에 설정 가능
    • 설정 변경 후 --save false로 미리보기 가능
    """
    # 입력 값 검증
    if legend_position and legend_position not in ["top", "bottom", "left", "right", "none"]:
        raise ValueError(f"잘못된 범례 위치: {legend_position}. 사용 가능한 위치: top, bottom, left, right, none")

    if data_label_position and data_label_position not in ["center", "above", "below", "left", "right", "outside", "inside"]:
        raise ValueError(
            f"잘못된 데이터 레이블 위치: {data_label_position}. 사용 가능한 위치: center, above, below, left, right, outside, inside"
        )

    if color_scheme and color_scheme not in ["colorful", "monochromatic", "office", "grayscale"]:
        raise ValueError(f"잘못된 색상 테마: {color_scheme}. 사용 가능한 테마: colorful, monochromatic, office, grayscale")

    if output_format not in ["json", "text"]:
        raise ValueError(f"잘못된 출력 형식: {output_format}. 사용 가능한 형식: json, text")

    book = None

    try:
        # 워크북 연결
        book = get_or_open_workbook(file_path=file_path, workbook_name=workbook_name, use_active=use_active, visible=visible)

        # 시트 가져오기
        target_sheet = get_sheet(book, sheet)

        # 차트 찾기
        chart = find_chart_by_name_or_index(target_sheet, chart_name, chart_index)

        # 설정 결과 추적
        configuration_results = {
            "chart_name": chart.name,
            "sheet": target_sheet.name,
            "applied_settings": {},
            "failed_settings": {},
            "platform": platform.system(),
        }

        # 차트 제목 설정
        if title:
            try:
                chart.api.HasTitle = True
                chart.api.ChartTitle.Text = title
                configuration_results["applied_settings"]["title"] = title
            except Exception as e:
                configuration_results["failed_settings"]["title"] = str(e)

        # 차트 스타일 설정
        if style:
            if set_chart_style(chart, style):
                configuration_results["applied_settings"]["style"] = style
            else:
                configuration_results["failed_settings"]["style"] = f"스타일 {style} 적용 실패 또는 지원되지 않음"

        # 범례 위치 설정
        if legend_position:
            if set_legend_position(chart, legend_position):
                configuration_results["applied_settings"]["legend_position"] = legend_position
            else:
                configuration_results["failed_settings"]["legend_position"] = "범례 위치 설정 실패"

        # 축 제목 설정
        if x_axis_title or y_axis_title:
            axis_results = set_axis_titles(chart, x_axis_title, y_axis_title)
            if x_axis_title:
                if axis_results["x_axis"]:
                    configuration_results["applied_settings"]["x_axis_title"] = x_axis_title
                else:
                    configuration_results["failed_settings"]["x_axis_title"] = "X축 제목 설정 실패"
            if y_axis_title:
                if axis_results["y_axis"]:
                    configuration_results["applied_settings"]["y_axis_title"] = y_axis_title
                else:
                    configuration_results["failed_settings"]["y_axis_title"] = "Y축 제목 설정 실패"

        # 데이터 레이블 설정
        if show_data_labels or hide_data_labels:
            show_labels = show_data_labels and not hide_data_labels
            if set_data_labels(chart, show_labels, data_label_position):
                configuration_results["applied_settings"]["data_labels"] = {
                    "show": show_labels,
                    "position": data_label_position if show_labels else None,
                }
            else:
                configuration_results["failed_settings"]["data_labels"] = "데이터 레이블 설정 실패 또는 지원되지 않음"

        # 색상 테마 설정
        if color_scheme:
            if set_chart_colors(chart, color_scheme):
                configuration_results["applied_settings"]["color_scheme"] = color_scheme
            else:
                configuration_results["failed_settings"]["color_scheme"] = "색상 테마 설정 실패 또는 지원되지 않음"

        # 배경 투명도 설정
        if transparent_bg:
            try:
                if platform.system() == "Windows":
                    chart.api.PlotArea.Format.Fill.Transparency = 1.0
                    chart.api.ChartArea.Format.Fill.Transparency = 1.0
                    configuration_results["applied_settings"]["transparent_background"] = True
                else:
                    configuration_results["failed_settings"]["transparent_background"] = "macOS에서는 지원되지 않음"
            except Exception as e:
                configuration_results["failed_settings"]["transparent_background"] = str(e)

        # 파일 저장
        if save and file_path:
            book.save()
            configuration_results["file_saved"] = True

        # 응답 생성
        applied_count = len(configuration_results["applied_settings"])
        failed_count = len(configuration_results["failed_settings"])

        message = f"차트 '{chart.name}' 설정 완료: {applied_count}개 적용"
        if failed_count > 0:
            message += f", {failed_count}개 실패"

        response = create_success_response(data=configuration_results, command="chart-configure", message=message)

        if output_format == "json":
            print(json.dumps(response, ensure_ascii=False, indent=2))
        else:
            # 텍스트 형식 출력
            print(f"=== 차트 설정 결과 ===")
            print(f"차트: {chart.name}")
            print(f"시트: {target_sheet.name}")
            print(f"플랫폼: {platform.system()}")
            print()

            if configuration_results["applied_settings"]:
                print("✅ 적용된 설정:")
                for setting, value in configuration_results["applied_settings"].items():
                    print(f"   {setting}: {value}")
                print()

            if configuration_results["failed_settings"]:
                print("❌ 실패한 설정:")
                for setting, error in configuration_results["failed_settings"].items():
                    print(f"   {setting}: {error}")
                print()

            if save and file_path:
                print("💾 파일이 저장되었습니다.")

    except Exception as e:
        error_response = create_error_response(e, "chart-configure")
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
    chart_configure()
