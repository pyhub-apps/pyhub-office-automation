"""
피벗차트 생성 명령어
피벗테이블 기반 동적 차트 생성 기능
"""

import json
import platform
from pathlib import Path
from typing import Optional

import typer
import xlwings as xw

from pyhub_office_automation.version import get_version

from .utils import create_error_response, create_success_response, get_or_open_workbook, get_sheet, normalize_path


def find_pivot_table(sheet, pivot_name):
    """시트에서 피벗테이블 찾기"""
    if platform.system() != "Windows":
        raise RuntimeError("피벗차트 생성은 Windows에서만 지원됩니다")

    try:
        # xlwings를 통해 피벗테이블 찾기 시도
        for pivot_table in sheet.api.PivotTables():
            if pivot_table.Name == pivot_name:
                return pivot_table

        raise ValueError(f"피벗테이블 '{pivot_name}'을 찾을 수 없습니다")

    except Exception as e:
        if "피벗테이블" in str(e):
            raise
        else:
            raise RuntimeError(f"피벗테이블 검색 중 오류 발생: {str(e)}")


def list_pivot_tables(sheet):
    """시트의 모든 피벗테이블 목록 반환"""
    if platform.system() != "Windows":
        return []

    try:
        pivot_names = []
        for pivot_table in sheet.api.PivotTables():
            pivot_names.append(pivot_table.Name)
        return pivot_names
    except:
        return []


def get_pivot_chart_type_constant(chart_type: str):
    """피벗차트 타입에 해당하는 xlwings 상수를 반환"""
    # 피벗차트에 적합한 차트 타입들 (상수값 직접 사용)
    pivot_chart_types = {
        "column": 51,  # xlColumnClustered
        "column_clustered": 51,
        "column_stacked": 52,  # xlColumnStacked
        "column_stacked_100": 53,  # xlColumnStacked100
        "bar": 57,  # xlBarClustered
        "bar_clustered": 57,
        "bar_stacked": 58,  # xlBarStacked
        "bar_stacked_100": 59,  # xlBarStacked100
        "pie": 5,  # xlPie
        "doughnut": -4120,  # xlDoughnut
        "line": 4,  # xlLine
        "line_markers": 65,  # xlLineMarkers
        "area": 1,  # xlArea
        "area_stacked": 76,  # xlAreaStacked
    }

    chart_type_lower = chart_type.lower()
    if chart_type_lower not in pivot_chart_types:
        raise ValueError(f"피벗차트에서 지원되지 않는 차트 타입: {chart_type}")

    # xlwings 상수를 시도하고, 실패하면 숫자값 직접 사용
    try:
        from xlwings.constants import ChartType

        const_map = {
            51: "xlColumnClustered",
            52: "xlColumnStacked",
            53: "xlColumnStacked100",
            57: "xlBarClustered",
            58: "xlBarStacked",
            59: "xlBarStacked100",
            5: "xlPie",
            -4120: "xlDoughnut",
            4: "xlLine",
            65: "xlLineMarkers",
            1: "xlArea",
            76: "xlAreaStacked",
        }

        chart_type_value = pivot_chart_types[chart_type_lower]
        const_name = const_map.get(chart_type_value)

        if const_name and hasattr(ChartType, const_name):
            return getattr(ChartType, const_name)
        else:
            # 상수 이름이 없거나 접근할 수 없으면 숫자값 직접 반환
            return chart_type_value

    except ImportError:
        # 상수를 가져올 수 없으면 숫자값 직접 반환
        return pivot_chart_types[chart_type_lower]


def chart_pivot_create(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="피벗차트를 생성할 Excel 파일의 절대 경로"),
    use_active: bool = typer.Option(False, "--use-active", help="현재 활성 워크북 사용"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help='열린 워크북 이름으로 접근 (예: "Sales.xlsx")'),
    pivot_name: str = typer.Option(..., "--pivot-name", help="차트를 생성할 피벗테이블 이름"),
    chart_type: str = typer.Option(
        "column",
        "--chart-type",
        help="피벗차트 유형 (column/column_clustered/column_stacked/column_stacked_100/bar/bar_clustered/bar_stacked/bar_stacked_100/pie/doughnut/line/line_markers/area/area_stacked, 기본값: column)",
    ),
    title: Optional[str] = typer.Option(None, "--title", help="피벗차트 제목"),
    position: str = typer.Option("H1", "--position", help="피벗차트 생성 위치 (셀 주소, 기본값: H1)"),
    width: int = typer.Option(400, "--width", help="피벗차트 너비 (픽셀, 기본값: 400)"),
    height: int = typer.Option(300, "--height", help="피벗차트 높이 (픽셀, 기본값: 300)"),
    sheet: Optional[str] = typer.Option(
        None, "--sheet", help="피벗차트를 생성할 시트 이름 (지정하지 않으면 피벗테이블과 같은 시트)"
    ),
    style: Optional[int] = typer.Option(None, "--style", help="피벗차트 스타일 번호 (1-48)"),
    legend_position: Optional[str] = typer.Option(None, "--legend-position", help="범례 위치 (top/bottom/left/right/none)"),
    show_data_labels: bool = typer.Option(False, "--show-data-labels", help="데이터 레이블 표시"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택 (json/text)"),
    visible: bool = typer.Option(False, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부 (기본값: False)"),
    save: bool = typer.Option(True, "--save", help="생성 후 파일 저장 여부 (기본값: True)"),
):
    """
    피벗테이블을 기반으로 동적 피벗차트를 생성합니다. (Windows 전용)

    기존 피벗테이블의 데이터를 활용하여 차트를 생성하며, 피벗테이블의 필드 변경에 따라
    차트도 자동으로 업데이트되는 동적 차트입니다. 대용량 데이터 분석에 특히 유용합니다.

    === 워크북 접근 방법 ===
    - --file-path: 파일 경로로 워크북 열기
    - --use-active: 현재 활성 워크북 사용 (권장)
    - --workbook-name: 열린 워크북 이름으로 접근 (예: "Sales.xlsx")

    === 피벗테이블 지정 ===
    --pivot-name 옵션으로 기준 피벗테이블을 지정합니다:

    • 피벗테이블 이름 확인: Excel에서 피벗테이블 선택 → 피벗테이블 도구 → 분석 탭
    • 존재하지 않는 이름 지정 시 사용 가능한 피벗테이블 목록 표시
    • 여러 시트에 동일 이름 피벗테이블이 있으면 처음 발견된 것 사용

    === 피벗차트의 특징과 장점 ===
    ▶ 동적 업데이트:
      • 피벗테이블 필터 변경 시 차트 자동 반영
      • 행/열 필드 재배치 시 차트 구조 자동 변경
      • 새 데이터 추가 후 피벗테이블 새로고침 시 차트도 업데이트

    ▶ 대용량 데이터 처리:
      • 수만 건 이상의 데이터도 빠른 차트 생성
      • 메모리 효율적인 데이터 처리
      • 필터링된 데이터만 차트에 반영

    === 지원되는 차트 유형 ===
    • column/column_clustered: 세로 막대 차트 (기본값, 시계열 데이터에 적합)
    • bar/bar_clustered: 가로 막대 차트 (카테고리 비교에 적합)
    • pie: 원형 차트 (구성 비율 표시에 적합)
    • line: 선 차트 (추세 분석에 적합)
    • area: 영역 차트 (누적 데이터 표시에 적합)

    === 위치 및 스타일 옵션 ===
    • --position "H1": 차트 생성 위치 (셀 주소)
    • --sheet "Charts": 차트 생성 대상 시트 (없으면 자동 생성)
    • --width 400 --height 300: 차트 크기
    • --style 1-48: 차트 스타일 (Windows)
    • --legend-position: 범례 위치
    • --show-data-labels: 데이터 레이블 표시

    === 실제 활용 시나리오 예제 ===

    # 1. 기본 피벗차트 생성
    oa excel chart-pivot-create --use-active --pivot-name "SalesAnalysis" --chart-type "column"

    # 2. 제품별 매출 비중 원형 차트
    oa excel chart-pivot-create --file-path "report.xlsx" --pivot-name "ProductSummary" \\
        --chart-type "pie" --title "제품별 판매 비중" --show-data-labels

    # 3. 지역별 매출 추세 분석
    oa excel chart-pivot-create --workbook-name "Dashboard.xlsx" --pivot-name "RegionalSales" \\
        --chart-type "line" --position "F5" --title "지역별 월간 매출 추세"

    # 4. 차트 전용 시트에 생성
    oa excel chart-pivot-create --use-active --pivot-name "QuarterlySummary" \\
        --chart-type "column" --sheet "피벗차트" --position "B2" --width 600 --height 400

    # 5. 스타일이 적용된 고급 피벗차트
    oa excel chart-pivot-create --file-path "analysis.xlsx" --pivot-name "YearlyTrend" \\
        --chart-type "area" --style 25 --legend-position "top" --title "연도별 매출 추이"

    === Windows 전용 기능 안내 ===
    • 이 명령어는 Windows에서만 작동합니다
    • macOS에서는 수동으로 피벗차트를 생성해주세요
    • COM API를 사용하여 Excel과 직접 연동

    === 피벗차트 활용 팁 ===
    • 피벗테이블 설계 시 차트 용도를 고려하여 필드 배치
    • 슬라이서 추가로 동적 필터링 기능 강화
    • 여러 피벗차트를 하나의 피벗테이블에서 생성하여 다각도 분석
    • 정기 보고서는 피벗차트로 구성하여 자동 업데이트 활용
    """
    # 입력 값 검증
    valid_chart_types = [
        "column",
        "column_clustered",
        "column_stacked",
        "column_stacked_100",
        "bar",
        "bar_clustered",
        "bar_stacked",
        "bar_stacked_100",
        "pie",
        "doughnut",
        "line",
        "line_markers",
        "area",
        "area_stacked",
    ]
    if chart_type not in valid_chart_types:
        raise ValueError(f"잘못된 차트 유형: {chart_type}. 사용 가능한 유형: {', '.join(valid_chart_types)}")

    if legend_position and legend_position not in ["top", "bottom", "left", "right", "none"]:
        raise ValueError(f"잘못된 범례 위치: {legend_position}. 사용 가능한 위치: top, bottom, left, right, none")

    if output_format not in ["json", "text"]:
        raise ValueError(f"잘못된 출력 형식: {output_format}. 사용 가능한 형식: json, text")

    book = None

    try:
        # Windows 전용 기능 확인
        if platform.system() != "Windows":
            raise RuntimeError("피벗차트 생성은 Windows에서만 지원됩니다. macOS에서는 수동으로 피벗차트를 생성해주세요.")

        # 워크북 연결
        book = get_or_open_workbook(file_path=file_path, workbook_name=workbook_name, use_active=use_active, visible=visible)

        # 피벗테이블이 있는 시트 찾기
        pivot_table = None
        pivot_sheet = None

        # 모든 시트에서 피벗테이블 검색
        for worksheet in book.sheets:
            try:
                pivot_table = find_pivot_table(worksheet, pivot_name)
                pivot_sheet = worksheet
                break
            except ValueError:
                continue  # 이 시트에는 해당 피벗테이블이 없음
            except Exception:
                continue  # 시트 검색 중 오류 발생, 다음 시트로

        if not pivot_table:
            # 사용 가능한 피벗테이블 목록 제공
            available_pivots = []
            for worksheet in book.sheets:
                pivot_names = list_pivot_tables(worksheet)
                for name in pivot_names:
                    available_pivots.append(f"{worksheet.name}!{name}")

            error_msg = f"피벗테이블 '{pivot_name}'을 찾을 수 없습니다."
            if available_pivots:
                error_msg += f" 사용 가능한 피벗테이블: {', '.join(available_pivots)}"
            else:
                error_msg += " 워크북에 피벗테이블이 없습니다."

            raise ValueError(error_msg)

        # 피벗차트 생성 대상 시트 결정
        if sheet:
            try:
                target_sheet = get_sheet(book, sheet)
            except ValueError:
                # 지정한 시트가 없으면 새로 생성
                target_sheet = book.sheets.add(name=sheet)
        else:
            # 시트가 지정되지 않으면 피벗테이블과 같은 시트 사용
            target_sheet = pivot_sheet

        # 차트 생성 위치 결정
        try:
            position_range = target_sheet.range(position)
            left = position_range.left
            top = position_range.top
        except Exception:
            # 잘못된 위치가 지정된 경우 기본 위치 사용
            left = 400
            top = 50

        # 차트 타입 상수 가져오기
        try:
            chart_type_const = get_pivot_chart_type_constant(chart_type)
        except Exception as e:
            raise ValueError(f"피벗차트 타입 처리 오류: {str(e)}")

        # 피벗차트 생성
        try:
            # 피벗차트 생성을 위한 COM API 사용
            chart_objects = target_sheet.api.ChartObjects()
            chart_object = chart_objects.Add(left, top, width, height)
            chart = chart_object.Chart

            # 피벗테이블을 소스로 설정
            chart.SetSourceData(pivot_table.TableRange1)
            chart.ChartType = chart_type_const

            # 피벗차트로 변경
            chart.PivotLayout.PivotTable = pivot_table

            chart_name = chart_object.Name

        except Exception as e:
            raise RuntimeError(f"피벗차트 생성 실패: {str(e)}")

        # 차트 제목 설정
        if title:
            try:
                chart.HasTitle = True
                chart.ChartTitle.Text = title
            except:
                pass

        # 차트 스타일 설정
        if style and 1 <= style <= 48:
            try:
                chart.ChartStyle = style
            except:
                pass

        # 범례 위치 설정
        if legend_position:
            try:
                if legend_position == "none":
                    chart.HasLegend = False
                else:
                    chart.HasLegend = True
                    from xlwings.constants import LegendPosition

                    legend_map = {
                        "top": LegendPosition.xlLegendPositionTop,
                        "bottom": LegendPosition.xlLegendPositionBottom,
                        "left": LegendPosition.xlLegendPositionLeft,
                        "right": LegendPosition.xlLegendPositionRight,
                    }
                    if legend_position in legend_map:
                        chart.Legend.Position = legend_map[legend_position]
            except:
                pass

        # 데이터 레이블 표시
        if show_data_labels:
            try:
                chart.FullSeriesCollection(1).HasDataLabels = True
            except:
                pass

        # 파일 저장
        if save and file_path:
            book.save()

        # 성공 응답 생성
        response_data = {
            "pivot_chart_name": chart_name,
            "pivot_table_name": pivot_name,
            "chart_type": chart_type,
            "source_sheet": pivot_sheet.name,
            "target_sheet": target_sheet.name,
            "position": position,
            "dimensions": {"width": width, "height": height},
            "workbook": book.name,
            "is_dynamic": True,
            "platform": "Windows",
        }

        if title:
            response_data["title"] = title

        response = create_success_response(
            data=response_data, command="chart-pivot-create", message=f"피벗차트 '{chart_name}'이 성공적으로 생성되었습니다"
        )

        if output_format == "json":
            print(json.dumps(response, ensure_ascii=False, indent=2))
        else:
            # 텍스트 형식 출력
            print(f"=== 피벗차트 생성 결과 ===")
            print(f"피벗차트: {chart_name}")
            print(f"피벗테이블: {pivot_name}")
            print(f"차트 유형: {chart_type}")
            print(f"소스 시트: {pivot_sheet.name}")
            print(f"대상 시트: {target_sheet.name}")
            print(f"위치: {position}")
            print(f"크기: {width} x {height}")
            if title:
                print(f"제목: {title}")
            print(f"\n✅ 동적 피벗차트가 생성되어 피벗테이블 변경 시 자동 업데이트됩니다.")
            if save and file_path:
                print("💾 파일이 저장되었습니다.")

    except Exception as e:
        error_response = create_error_response(e, "chart-pivot-create")
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
    chart_pivot_create()
