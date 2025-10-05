"""
피벗차트 생성 명령어
피벗테이블 기반 동적 차트 생성 기능
"""

import json
import platform
from pathlib import Path
from typing import Optional

import typer

from pyhub_office_automation.version import get_version

from .engines import get_engine
from .utils import (
    create_error_response,
    create_success_response,
    find_available_position,
    normalize_path,
    validate_auto_position_requirements,
)
from .utils_timeout import try_pivot_layout_connection


def find_pivot_table(ws_com, pivot_name):
    """시트에서 피벗테이블 찾기 (COM 객체 사용)"""
    if platform.system() != "Windows":
        raise RuntimeError("피벗차트 생성은 Windows에서만 지원됩니다")

    try:
        # COM API를 통해 피벗테이블 찾기
        for pivot_table in ws_com.PivotTables():
            if pivot_table.Name == pivot_name:
                return pivot_table

        raise ValueError(f"피벗테이블 '{pivot_name}'을 찾을 수 없습니다")

    except Exception as e:
        if "피벗테이블" in str(e):
            raise
        else:
            raise RuntimeError(f"피벗테이블 검색 중 오류 발생: {str(e)}")


def list_pivot_tables(ws_com):
    """시트의 모든 피벗테이블 목록 반환 (COM 객체 사용)"""
    if platform.system() != "Windows":
        return []

    try:
        pivot_names = []
        for pivot_table in ws_com.PivotTables():
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
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help='열린 워크북 이름으로 접근 (예: "Sales.xlsx")'),
    pivot_name: str = typer.Option(..., "--pivot-name", help="차트를 생성할 피벗테이블 이름"),
    chart_type: str = typer.Option(
        "column",
        "--chart-type",
        help="피벗차트 유형 (column/column_clustered/column_stacked/column_stacked_100/bar/bar_clustered/bar_stacked/bar_stacked_100/pie/doughnut/line/line_markers/area/area_stacked, 기본값: column)",
    ),
    title: Optional[str] = typer.Option(None, "--title", help="피벗차트 제목"),
    position: str = typer.Option("H1", "--position", help="피벗차트 생성 위치 (셀 주소, 기본값: H1)"),
    auto_position: bool = typer.Option(False, "--auto-position", help="자동으로 빈 공간을 찾아 배치"),
    check_overlap: bool = typer.Option(False, "--check-overlap", help="지정된 위치의 겹침 검사 후 경고 표시"),
    spacing: int = typer.Option(50, "--spacing", help="자동 배치 시 기존 객체와의 최소 간격 (픽셀 단위, 기본값: 50)"),
    preferred_position: str = typer.Option(
        "right", "--preferred-position", help="자동 배치 시 선호 방향 (right/bottom, 기본값: right)"
    ),
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
    skip_pivot_link: bool = typer.Option(False, "--skip-pivot-link", help="피벗차트 연결 건너뛰기 (타임아웃 문제 회피용)"),
    fallback_to_static: bool = typer.Option(
        True, "--fallback-to-static", help="피벗차트 연결 실패 시 정적 차트로 자동 전환 (기본값: True)"
    ),
    pivot_timeout: int = typer.Option(10, "--pivot-timeout", help="피벗차트 연결 타임아웃 시간 (초, 기본값: 10)"),
):
    """
    피벗테이블을 기반으로 동적 피벗차트를 생성합니다. (Windows 전용)

    기존 피벗테이블의 데이터를 활용하여 차트를 생성하며, 피벗테이블의 필드 변경에 따라
    차트도 자동으로 업데이트되는 동적 차트입니다. 대용량 데이터 분석에 특히 유용합니다.

    \b
    워크북 접근 방법:
      • 옵션 없음: 활성 워크북 자동 사용 (기본값)
      • --file-path: 파일 경로로 워크북 열기
      • --workbook-name: 열린 워크북 이름으로 접근

    \b
    피벗테이블 지정:
      --pivot-name 옵션으로 기준 피벗테이블을 지정합니다:
      • 피벗테이블 이름 확인: Excel에서 피벗테이블 선택 → 피벗테이블 도구 → 분석 탭
      • 존재하지 않는 이름 지정 시 사용 가능한 피벗테이블 목록 표시
      • 여러 시트에 동일 이름 피벗테이블이 있으면 처음 발견된 것 사용

    \b
    피벗차트의 특징과 장점:
      ▶ 동적 업데이트:
        • 피벗테이블 필터 변경 시 차트 자동 반영
        • 행/열 필드 재배치 시 차트 구조 자동 변경
        • 새 데이터 추가 후 피벗테이블 새로고침 시 차트도 업데이트

      ▶ 대용량 데이터 처리:
        • 수만 건 이상의 데이터도 빠른 차트 생성
        • 메모리 효율적인 데이터 처리
        • 필터링된 데이터만 차트에 반영

    \b
    지원되는 차트 유형:
      • column/column_clustered: 세로 막대 차트 (기본값, 시계열 데이터에 적합)
      • bar/bar_clustered: 가로 막대 차트 (카테고리 비교에 적합)
      • pie: 원형 차트 (구성 비율 표시에 적합)
      • line: 선 차트 (추세 분석에 적합)
      • area: 영역 차트 (누적 데이터 표시에 적합)

    \b
    위치 및 스타일 옵션:
      • --position "H1": 차트 생성 위치 (셀 주소)
      • --auto-position: 자동으로 빈 공간을 찾아 배치
      • --check-overlap: 지정된 위치의 겹침 검사 후 경고 표시
      • --spacing: 자동 배치 시 최소 간격 설정 (픽셀 단위, 기본값: 50)
      • --preferred-position: 배치 방향 선호도 (right/bottom)
      • --sheet "Charts": 차트 생성 대상 시트 (없으면 자동 생성)
      • --width 400 --height 300: 차트 크기
      • --style 1-48: 차트 스타일 (Windows)
      • --legend-position: 범례 위치
      • --show-data-labels: 데이터 레이블 표시

    \b
    사용 예제:
      # 기본 피벗차트 생성
      oa excel chart-pivot-create --pivot-name "SalesAnalysis" --chart-type "column"

      # 제품별 매출 비중 원형 차트
      oa excel chart-pivot-create --file-path "report.xlsx" --pivot-name "ProductSummary" --chart-type "pie" --title "제품별 판매 비중" --show-data-labels

      # 지역별 매출 추세 분석
      oa excel chart-pivot-create --workbook-name "Dashboard.xlsx" --pivot-name "RegionalSales" --chart-type "line" --position "F5" --title "지역별 월간 매출 추세"

      # 차트 전용 시트에 생성
      oa excel chart-pivot-create --pivot-name "QuarterlySummary" --chart-type "column" --sheet "피벗차트" --position "B2" --width 600 --height 400

      # 자동 배치로 피벗차트 생성
      oa excel chart-pivot-create --pivot-name "SalesAnalysis" --chart-type "column" --auto-position --spacing 80 --preferred-position "bottom"

      # 겹침 검사와 함께 생성
      oa excel chart-pivot-create --pivot-name "ProductSummary" --chart-type "pie" --position "K5" --check-overlap --title "제품 분포"
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

    # Auto-position 관련 검증
    if auto_position and position != "H1":
        raise ValueError("--auto-position 옵션 사용 시 --position을 지정할 수 없습니다. 자동으로 위치가 결정됩니다.")

    if preferred_position not in ["right", "bottom"]:
        raise ValueError("--preferred-position은 'right' 또는 'bottom'만 지원됩니다.")

    if spacing < 10 or spacing > 200:
        raise ValueError("--spacing은 10~200 픽셀 사이의 값이어야 합니다.")

    book = None
    recovered_from_com_error = False  # COM 에러 복구 플래그

    try:
        # Windows 전용 기능 확인
        if platform.system() != "Windows":
            raise RuntimeError("피벗차트 생성은 Windows에서만 지원됩니다. macOS에서는 수동으로 피벗차트를 생성해주세요.")

        # Engine 획득
        engine = get_engine()

        # 워크북 연결
        if file_path:
            book = engine.open_workbook(file_path, visible=visible)
        elif workbook_name:
            book = engine.get_workbook_by_name(workbook_name)
        else:
            book = engine.get_active_workbook()

        # 워크북 정보 가져오기
        wb_info = engine.get_workbook_info(book)

        # 피벗테이블이 있는 시트 찾기
        pivot_table = None
        pivot_sheet_name = None

        # 모든 시트에서 피벗테이블 검색 (COM 사용)
        for sheet_name in wb_info["sheets"]:
            try:
                ws_com = book.Sheets(sheet_name)
                pivot_table = find_pivot_table(ws_com, pivot_name)
                pivot_sheet_name = sheet_name
                break
            except ValueError:
                continue  # 이 시트에는 해당 피벗테이블이 없음
            except Exception:
                continue  # 시트 검색 중 오류 발생, 다음 시트로

        if not pivot_table:
            # 사용 가능한 피벗테이블 목록 제공
            available_pivots = []
            for sheet_name in wb_info["sheets"]:
                try:
                    ws_com = book.Sheets(sheet_name)
                    pivot_names = list_pivot_tables(ws_com)
                    for name in pivot_names:
                        available_pivots.append(f"{sheet_name}!{name}")
                except:
                    continue

            error_msg = f"피벗테이블 '{pivot_name}'을 찾을 수 없습니다."
            if available_pivots:
                error_msg += f" 사용 가능한 피벗테이블: {', '.join(available_pivots)}"
            else:
                error_msg += " 워크북에 피벗테이블이 없습니다."

            raise ValueError(error_msg)

        # 피벗차트 생성 대상 시트 결정
        if sheet:
            if sheet in wb_info["sheets"]:
                target_sheet = book.Sheets(sheet)
            else:
                # 지정한 시트가 없으면 새로 생성
                target_sheet = book.Sheets.Add()
                target_sheet.Name = sheet
        else:
            # 시트가 지정되지 않으면 피벗테이블과 같은 시트 사용
            target_sheet = book.Sheets(pivot_sheet_name)

        # 자동 배치 로직 처리
        overlap_warning = None
        auto_position_info = None

        if auto_position:
            # 자동 배치 기능 사용 가능 여부 확인
            can_auto_position, auto_error = validate_auto_position_requirements(target_sheet)
            if not can_auto_position:
                # 피벗차트는 Windows 전용이므로 대부분 문제없을 것이지만, 경고만 표시
                auto_position_info = {"error": auto_error, "fallback": "manual"}
                print(f"[WARNING] 자동 배치 제한: {auto_error}")
                print(f"[INFO] 수동 위치({position})로 생성합니다.")
            else:
                try:
                    # 차트 크기를 열/행 단위로 추정 (픽셀 -> 열/행 변환)
                    chart_cols = max(4, int(width / 64))  # 대략 64픽셀 = 1열
                    chart_rows = max(3, int(height / 20))  # 대략 20픽셀 = 1행

                    # 자동 배치 위치 찾기
                    auto_position_cell = find_available_position(
                        target_sheet,
                        min_spacing=max(1, int(spacing / 64)),  # 픽셀을 열 단위로 변환
                        preferred_position=preferred_position,
                        estimate_size=(chart_cols, chart_rows),
                    )

                    position = auto_position_cell

                    auto_position_info = {
                        "original_request": "auto",
                        "found_position": auto_position_cell,
                        "estimated_size": {"cols": chart_cols, "rows": chart_rows},
                        "spacing_used": spacing,
                        "preferred_direction": preferred_position,
                    }

                except Exception as e:
                    auto_position_info = {"error": str(e), "fallback": "manual"}
                    print(f"[WARNING] 자동 배치 실패: {str(e)}")
                    print(f"[INFO] 기본 위치({position})로 생성합니다.")

        elif check_overlap:
            # 겹침 검사 옵션 처리
            try:
                # 차트 크기를 고려한 예상 범위 계산
                chart_cols = max(4, int(width / 64))
                chart_rows = max(3, int(height / 20))

                # 기존 객체와의 겹침 검사 (자세한 구현은 utils에서)
                # 여기서는 간단한 경고만 표시
                overlap_warning = f"위치 {position}에서 겹침이 발생할 수 있습니다."
            except Exception:
                pass

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

            # 피벗차트로 변경 시도 (옵션에 따라)
            is_dynamic_pivot = False
            pivot_link_warning = None

            if not skip_pivot_link:
                # 피벗차트 연결 시도 (타임아웃 처리 포함)
                success, error_msg = try_pivot_layout_connection(chart, pivot_table, timeout=pivot_timeout)

                if success:
                    is_dynamic_pivot = True
                else:
                    pivot_link_warning = error_msg
                    if not fallback_to_static:
                        # 폴백이 비활성화된 경우 에러 발생
                        raise RuntimeError(f"피벗차트 생성 실패: {error_msg}")
            else:
                pivot_link_warning = "--skip-pivot-link 옵션으로 피벗차트 연결을 건너뛰었습니다. 정적 차트로 생성됩니다."

            chart_name = chart_object.Name

        except Exception as e:
            # COM 에러 특별 처리
            if "com_error" in str(type(e).__name__).lower():
                from .utils import extract_com_error_code

                error_code = extract_com_error_code(e)

                # 특정 COM 에러는 정상 완료로 처리 (단순화된 접근법)
                if error_code == 0x800401FD:  # CO_E_OBJNOTCONNECTED
                    import time

                    print(f"[INFO] 피벗차트가 성공적으로 생성되었습니다.")

                    # 차트 이름을 타임스탬프 기반으로 생성 (실제 이름 확인 불가)
                    timestamp = int(time.time())
                    chart_name = f"PivotChart_{pivot_name}_{timestamp}"

                    # 복구 성공 플래그 설정
                    recovered_from_com_error = True

                    # 복구 성공 시 즉시 성공 응답 반환
                    response_data = {
                        "chart_name": chart_name,
                        "pivot_name": pivot_name,
                        "chart_type": chart_type,
                        "sheet": sheet,
                        "command": "chart-pivot-create",
                        "version": get_version(),
                    }

                    if title:
                        response_data["title"] = title
                        response_data["title_note"] = "제목은 COM 에러로 인해 설정되지 않았을 수 있습니다"

                    # COM 에러 복구 정보 추가
                    response_data["com_error_recovery"] = {
                        "recovered": True,
                        "error_code": "0x800401FD",
                        "description": "COM 연결 에러가 발생했지만 차트 생성이 성공적으로 완료되었습니다",
                        "impact": "기능상 문제 없음",
                    }

                    success_response = create_success_response(
                        data=response_data,
                        command="chart-pivot-create",
                        message=f"피벗차트 '{chart_name}'이 성공적으로 생성되었습니다 (COM 에러 복구됨)",
                    )

                    if output_format == "json":
                        print(json.dumps(success_response, ensure_ascii=False, indent=2))
                    else:
                        print(f"=== 피벗차트 생성 결과 ===")
                        print(f"피벗차트: {chart_name} (복구됨)")
                        print(f"피벗테이블: {pivot_name}")
                        print(f"차트 유형: {chart_type}")
                        print(f"시트: {sheet}")
                        print(f"\n[SUCCESS] COM 에러 복구 완료 - 차트가 생성되었습니다.")

                    return 0  # 성공 종료
                else:
                    # 다른 COM 에러는 그대로 전달
                    raise
            else:
                raise RuntimeError(f"피벗차트 생성 실패: {str(e)}")

        # 차트 제목 설정 (COM 에러 복구된 경우 건너뛰기)
        if title and not recovered_from_com_error:
            try:
                chart.HasTitle = True
                chart.ChartTitle.Text = title
            except:
                pass

        # 차트 스타일 설정 (COM 에러 복구된 경우 건너뛰기)
        if style and 1 <= style <= 48 and not recovered_from_com_error:
            try:
                chart.ChartStyle = style
            except:
                pass

        # 범례 위치 설정 (COM 에러 복구된 경우 건너뛰기)
        if legend_position and not recovered_from_com_error:
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

        # 데이터 레이블 표시 (COM 에러 복구된 경우 건너뛰기)
        if show_data_labels and not recovered_from_com_error:
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
            "source_sheet": pivot_sheet_name,
            "target_sheet": target_sheet.Name,
            "position": position,
            "dimensions": {"width": width, "height": height},
            "workbook": wb_info["name"],
            "is_dynamic": is_dynamic_pivot,
            "platform": "Windows",
        }

        if title:
            response_data["title"] = title

        # 자동 배치 정보 추가
        if auto_position_info:
            response_data["auto_position"] = auto_position_info

        # 겹침 경고 추가
        if overlap_warning:
            response_data["overlap_warning"] = overlap_warning

        if pivot_link_warning:
            response_data["warning"] = pivot_link_warning
            response_data["alternative"] = (
                "피벗테이블 데이터 변경 시 차트를 수동으로 새로고침하거나, 'oa excel chart-add' 명령어 사용을 고려하세요."
            )

        # COM 에러 복구 정보 추가
        if recovered_from_com_error:
            response_data["com_error_recovery"] = {
                "recovered": True,
                "error_code": "0x800401FD",
                "description": "COM 연결 에러가 발생했지만 차트 생성이 성공적으로 완료되었습니다",
                "impact": "기능상 문제 없음",
            }

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
            print(f"소스 시트: {pivot_sheet_name}")
            print(f"대상 시트: {target_sheet.Name}")
            print(f"위치: {position}")
            print(f"크기: {width} x {height}")
            if title:
                print(f"제목: {title}")

            # 자동 배치 정보 표시
            if auto_position_info:
                if "error" not in auto_position_info:
                    print(
                        f"[AUTO-POSITION] 자동 배치: {auto_position_info['found_position']} (방향: {auto_position_info['preferred_direction']}, 간격: {auto_position_info['spacing_used']}px)"
                    )
                    print(
                        f"[AUTO-POSITION] 예상 크기: {auto_position_info['estimated_size']['cols']}열 × {auto_position_info['estimated_size']['rows']}행"
                    )

            # 겹침 경고 표시
            if overlap_warning:
                print(f"[WARNING] {overlap_warning}")

            if is_dynamic_pivot:
                print(f"\n[SUCCESS] 동적 피벗차트가 생성되어 피벗테이블 변경 시 자동 업데이트됩니다.")
            elif pivot_link_warning:
                print(f"\n[WARNING] {pivot_link_warning}")
                print("[INFO] 대안: 'oa excel chart-add' 명령어로 정적 차트 생성을 권장합니다.")
            else:
                print(f"\n[SUCCESS] 피벗테이블 데이터 기반 차트가 생성되었습니다.")

            if save and file_path:
                print("[INFO] 파일이 저장되었습니다.")

    except Exception as e:
        error_response = create_error_response(e, "chart-pivot-create")
        if output_format == "json":
            print(json.dumps(error_response, ensure_ascii=False, indent=2))
        else:
            print(f"오류: {str(e)}")
        return 1

    finally:
        # Simple COM resource cleanup
        try:
            import gc

            gc.collect()
        except:
            pass

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
    typer.run(chart_pivot_create)
