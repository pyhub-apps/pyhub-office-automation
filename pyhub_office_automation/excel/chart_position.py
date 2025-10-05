"""
차트 위치 조정 명령어
차트의 위치와 크기를 정밀하게 조정하는 기능
"""

import json
import platform
from pathlib import Path
from typing import Optional

import typer

from pyhub_office_automation.version import get_version

from .engines import get_engine
from .utils import create_error_response, create_success_response, get_sheet


def find_chart_by_name_or_index(sheet, chart_name=None, chart_index=None):
    """차트 이름이나 인덱스로 차트 객체 찾기"""
    if platform.system() == "Windows":
        # Windows: COM API 사용
        chart_objects = sheet.ChartObjects()

        if chart_name:
            for i in range(1, chart_objects.Count + 1):
                chart_obj = chart_objects(i)
                if chart_obj.Name == chart_name:
                    return chart_obj
            raise ValueError(f"차트 '{chart_name}'을 찾을 수 없습니다")

        elif chart_index is not None:
            if 0 <= chart_index < chart_objects.Count:
                return chart_objects(chart_index + 1)  # COM is 1-indexed
            else:
                raise IndexError(f"차트 인덱스 {chart_index}는 범위를 벗어났습니다 (0-{chart_objects.Count-1})")
        else:
            raise ValueError("차트 이름(--chart-name) 또는 인덱스(--chart-index) 중 하나를 지정해야 합니다")
    else:
        # macOS: xlwings 방식
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


def get_cell_position(sheet, cell_address):
    """셀 주소에서 픽셀 위치 계산"""
    try:
        if platform.system() == "Windows":
            # Windows: COM API 사용
            cell_range = sheet.Range(cell_address)
            return {"left": cell_range.Left, "top": cell_range.Top, "width": cell_range.Width, "height": cell_range.Height}
        else:
            # macOS: xlwings 방식
            cell_range = sheet.range(cell_address)
            return {"left": cell_range.left, "top": cell_range.top, "width": cell_range.width, "height": cell_range.height}
    except Exception:
        raise ValueError(f"잘못된 셀 주소입니다: {cell_address}")


def find_shape_by_name(sheet, shape_name):
    """시트에서 도형 이름으로 도형 찾기"""
    try:
        if platform.system() == "Windows":
            # Windows: COM API 사용
            for shape in sheet.Shapes:
                if shape.Name == shape_name:
                    return {"left": shape.Left, "top": shape.Top, "width": shape.Width, "height": shape.Height}
        else:
            # macOS: xlwings 방식
            if hasattr(sheet, "shapes"):
                for shape in sheet.shapes:
                    if shape.name == shape_name:
                        return {"left": shape.left, "top": shape.top, "width": shape.width, "height": shape.height}

        raise ValueError(f"도형 '{shape_name}'을 찾을 수 없습니다")
    except Exception as e:
        raise ValueError(f"도형 검색 중 오류: {str(e)}")


def calculate_relative_position(base_position, relative_direction, offset=10):
    """기준 위치에서 상대 위치 계산"""
    directions = {
        "right": {"left": base_position["left"] + base_position["width"] + offset, "top": base_position["top"]},
        "left": {"left": base_position["left"] - offset, "top": base_position["top"]},
        "below": {"left": base_position["left"], "top": base_position["top"] + base_position["height"] + offset},
        "above": {"left": base_position["left"], "top": base_position["top"] - offset},
        "center": {
            "left": base_position["left"] + (base_position["width"] / 2),
            "top": base_position["top"] + (base_position["height"] / 2),
        },
    }

    if relative_direction not in directions:
        raise ValueError(f"지원되지 않는 상대 위치: {relative_direction}")

    return directions[relative_direction]


def chart_position(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="차트가 있는 Excel 파일의 절대 경로"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help='열린 워크북 이름으로 접근 (예: "Sales.xlsx")'),
    sheet: Optional[str] = typer.Option(None, "--sheet", help="차트가 있는 시트 이름 (지정하지 않으면 활성 시트)"),
    name: Optional[str] = typer.Option(None, "--name", help="조정할 차트의 이름"),
    chart_name: Optional[str] = typer.Option(None, "--chart-name", help="[별칭] 조정할 차트의 이름 (--name 사용 권장)"),
    index: Optional[int] = typer.Option(None, "--index", help="조정할 차트의 인덱스 (0부터 시작)"),
    chart_index: Optional[int] = typer.Option(None, "--chart-index", help="[별칭] 조정할 차트의 인덱스 (--index 사용 권장)"),
    left: Optional[float] = typer.Option(None, "--left", help="차트의 왼쪽 위치 (픽셀)"),
    top: Optional[float] = typer.Option(None, "--top", help="차트의 위쪽 위치 (픽셀)"),
    width: Optional[float] = typer.Option(None, "--width", help="차트의 너비 (픽셀)"),
    height: Optional[float] = typer.Option(None, "--height", help="차트의 높이 (픽셀)"),
    anchor_cell: Optional[str] = typer.Option(None, "--anchor-cell", help='차트를 고정할 셀 주소 (예: "E5")'),
    relative_to: Optional[str] = typer.Option(None, "--relative-to", help="상대 위치 기준이 될 도형 이름"),
    relative_direction: Optional[str] = typer.Option(
        None, "--relative-direction", help="상대 위치 방향 (right/left/below/above/center, --relative-to와 함께 사용)"
    ),
    offset: int = typer.Option(10, "--offset", help="상대 위치 설정 시 간격 (픽셀, 기본값: 10)"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택 (json/text)"),
    visible: bool = typer.Option(False, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부 (기본값: False)"),
    save: bool = typer.Option(True, "--save", help="조정 후 파일 저장 여부 (기본값: True)"),
):
    """
    차트의 위치와 크기를 정밀하게 조정합니다.

    생성된 차트의 위치를 셀 기준 또는 픽셀 단위로 이동시키고, 크기를 조정할 수 있습니다.
    대시보드 레이아웃 구성, 차트 정렬, 프레젠테이션 슬라이드 배치에 특히 유용합니다.

    === 워크북 접근 방법 ===
    - --file-path: 파일 경로로 워크북 열기
        - --workbook-name: 열린 워크북 이름으로 접근 (예: "Sales.xlsx")

    === 차트 선택 방법 ===
    대상 차트를 지정하는 두 가지 방법:

    ▶ 차트 이름으로 선택:
      • --chart-name "Chart1"
      • chart-list 명령으로 차트 이름 확인

    ▶ 인덱스 번호로 선택:
      • --chart-index 0 (첫 번째 차트)
      • 시트의 차트 순서대로 0, 1, 2...

    === 위치 조정 방법 ===

    ▶ 셀 기준 위치 설정 (권장):
      • --anchor-cell "E5": E5 셀 위치에 차트 고정
      • 가장 직관적이고 Excel 그리드에 맞춘 배치
      • 열/행 삽입 시에도 상대적 위치 유지

    ▶ 절대 픽셀 위치:
      • --left 300 --top 100: 픽셀 단위 정확한 위치
      • 정밀한 레이아웃이 필요한 경우 사용
      • 화면 해상도나 Excel 창 크기에 따라 달라질 수 있음

    ▶ 상대 위치 설정:
      • --relative-to "ChartBox1" --relative-direction "right"
      • 다른 도형이나 차트를 기준으로 상대적 배치
      • --offset 옵션으로 간격 조정 (기본 10px)

    === 크기 조정 방법 ===
    • --width 500: 차트 너비 (픽셀)
    • --height 300: 차트 높이 (픽셀)
    • 비율 유지 없이 자유롭게 조정 가능

    === 위치 설정 전략 ===

    ▶ 대시보드 레이아웃:
      • 셀 기준 위치 사용으로 일관된 배치
      • 표준 크기 설정 (예: 400x300, 500x350)
      • 행/열 간격을 고려한 규칙적 배치

    ▶ 프레젠테이션:
      • 절대 픽셀 위치로 정밀 배치
      • 화면 크기에 맞는 최적 크기 설정
      • 슬라이드 비율 고려 (16:9, 4:3)

    ▶ 인쇄용 레포트:
      • 셀 기준 위치로 페이지 레이아웃에 맞춤
      • A4 용지 기준 적절한 크기 설정
      • 여백과 텍스트 영역 고려

    === 실제 활용 시나리오 예제 ===

    # 1. 셀 기준 차트 이동 (가장 일반적)
    oa excel chart-position --chart-index 0 --anchor-cell "H2"

    # 2. 차트 크기 조정과 위치 이동
    oa excel chart-position --chart-name "SalesChart" \\
        --anchor-cell "B10" --width 600 --height 400

    # 3. 픽셀 단위 정밀 배치 (프레젠테이션용)
    oa excel chart-position --file-path "presentation.xlsx" --chart-index 0 \\
        --left 50 --top 80 --width 800 --height 500

    # 4. 여러 차트 정렬 배치
    oa excel chart-position --chart-index 0 --anchor-cell "B2"
    oa excel chart-position --chart-index 1 --anchor-cell "H2"
    oa excel chart-position --chart-index 2 --anchor-cell "B15"
    oa excel chart-position --chart-index 3 --anchor-cell "H15"

    # 5. 상대 위치 기반 배치 (고급)
    oa excel chart-position --workbook-name "Dashboard.xlsx" --chart-name "Chart2" \\
        --relative-to "Chart1" --relative-direction "right" --offset 20

    # 6. 대시보드 표준 크기로 일괄 조정
    oa excel chart-position --chart-index 0 --width 450 --height 300
    oa excel chart-position --chart-index 1 --width 450 --height 300

    === 차트 배치 모범 사례 ===
    • 일관된 크기: 같은 대시보드 내 차트는 동일하거나 조화로운 크기 사용
    • 규칙적 간격: 차트 간 일정한 간격으로 시각적 안정감 조성
    • 데이터 중요도 반영: 중요한 차트는 상단 또는 좌측에 배치
    • 읽기 흐름 고려: 왼쪽 → 오른쪽, 위 → 아래 순서로 배치

    === 팁 ===
    • chart-list --detailed로 현재 위치/크기 확인 후 조정
    • --save false로 미리보기 후 최종 적용
    • 여러 차트 일괄 조정 시 스크립트나 반복 명령 활용
    """
    # 입력 값 검증
    if relative_direction and relative_direction not in ["right", "left", "below", "above", "center"]:
        raise ValueError(f"잘못된 상대 위치 방향: {relative_direction}. 사용 가능한 방향: right, left, below, above, center")

    fmt = str(output_format) if output_format else "json"
    if fmt not in ["json", "text"]:
        raise ValueError(f"잘못된 출력 형식: {fmt}. 사용 가능한 형식: json, text")

    try:
        # 옵션 우선순위 처리 (새 옵션 우선)
        target_name = name or chart_name
        target_index = index if index is not None else chart_index

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

        # 시트 가져오기 (COM API 직접 사용)
        if platform.system() == "Windows":
            if sheet:
                target_sheet = book.Sheets(sheet)
            else:
                target_sheet = book.ActiveSheet
        else:
            # macOS는 xlwings 방식 유지
            import xlwings as xw

            xw_book = xw.books[wb_info["name"]]
            target_sheet = get_sheet(xw_book, sheet)

        # 차트 찾기
        chart = find_chart_by_name_or_index(target_sheet, target_name, target_index)

        # 현재 차트 위치 및 크기 저장 (플랫폼별)
        if platform.system() == "Windows":
            original_position = {"left": chart.Left, "top": chart.Top, "width": chart.Width, "height": chart.Height}
        else:
            original_position = {"left": chart.left, "top": chart.top, "width": chart.width, "height": chart.height}

        # 새로운 위치 계산
        new_position = {"left": None, "top": None}
        new_size = {"width": None, "height": None}

        # 위치 설정 우선순위: 1) 상대 위치 2) 셀 기준 3) 절대 위치
        if relative_to and relative_direction:
            # 상대 위치 설정
            base_position = find_shape_by_name(target_sheet, relative_to)
            relative_pos = calculate_relative_position(base_position, relative_direction, offset)
            new_position["left"] = relative_pos["left"]
            new_position["top"] = relative_pos["top"]

        elif anchor_cell:
            # 셀 기준 위치 설정
            cell_pos = get_cell_position(target_sheet, anchor_cell)
            new_position["left"] = cell_pos["left"]
            new_position["top"] = cell_pos["top"]

        else:
            # 절대 위치 설정
            if left is not None:
                new_position["left"] = left
            if top is not None:
                new_position["top"] = top

        # 크기 설정
        if width is not None:
            new_size["width"] = width
        if height is not None:
            new_size["height"] = height

        # 변경사항 추적
        changes_made = {}
        position_changed = False
        size_changed = False

        # 위치 적용 (플랫폼별)
        if platform.system() == "Windows":
            if new_position["left"] is not None:
                chart.Left = new_position["left"]
                changes_made["left"] = new_position["left"]
                position_changed = True

            if new_position["top"] is not None:
                chart.Top = new_position["top"]
                changes_made["top"] = new_position["top"]
                position_changed = True

            # 크기 적용
            if new_size["width"] is not None:
                chart.Width = new_size["width"]
                changes_made["width"] = new_size["width"]
                size_changed = True

            if new_size["height"] is not None:
                chart.Height = new_size["height"]
                changes_made["height"] = new_size["height"]
                size_changed = True
        else:
            # macOS: xlwings 방식
            if new_position["left"] is not None:
                chart.left = new_position["left"]
                changes_made["left"] = new_position["left"]
                position_changed = True

            if new_position["top"] is not None:
                chart.top = new_position["top"]
                changes_made["top"] = new_position["top"]
                position_changed = True

            if new_size["width"] is not None:
                chart.width = new_size["width"]
                changes_made["width"] = new_size["width"]
                size_changed = True

            if new_size["height"] is not None:
                chart.height = new_size["height"]
                changes_made["height"] = new_size["height"]
                size_changed = True

        # 변경사항이 없는 경우 확인
        if not changes_made:
            raise ValueError("변경할 위치나 크기 정보가 지정되지 않았습니다")

        # 파일 저장
        if save and file_path:
            if platform.system() == "Windows":
                book.Save()
            else:
                book.save()

        # 차트 이름 및 시트 이름 가져오기 (플랫폼별)
        if platform.system() == "Windows":
            chart_name_str = chart.Name
            sheet_name_str = target_sheet.Name
            current_pos = {"left": chart.Left, "top": chart.Top, "width": chart.Width, "height": chart.Height}
        else:
            chart_name_str = chart.name
            sheet_name_str = target_sheet.name
            current_pos = {"left": chart.left, "top": chart.top, "width": chart.width, "height": chart.height}

        # 응답 데이터 구성
        response_data = {
            "chart_name": chart_name_str,
            "sheet": sheet_name_str,
            "original_position": original_position,
            "new_position": current_pos,
            "changes_applied": changes_made,
            "position_changed": position_changed,
            "size_changed": size_changed,
        }

        # 설정 방법 정보 추가
        if relative_to and relative_direction:
            response_data["positioning_method"] = {
                "type": "relative",
                "relative_to": relative_to,
                "direction": relative_direction,
                "offset": offset,
            }
        elif anchor_cell:
            response_data["positioning_method"] = {"type": "cell_anchor", "anchor_cell": anchor_cell}
        else:
            response_data["positioning_method"] = {"type": "absolute", "coordinates": changes_made}

        if save and file_path:
            response_data["file_saved"] = True

        message = f"차트 '{chart_name_str}' 위치/크기 조정 완료"
        if position_changed and size_changed:
            message += " (위치 및 크기 변경)"
        elif position_changed:
            message += " (위치 변경)"
        elif size_changed:
            message += " (크기 변경)"

        response = create_success_response(data=response_data, command="chart-position", message=message)

        if fmt == "json":
            print(json.dumps(response, ensure_ascii=False, indent=2))
        else:
            # 텍스트 형식 출력
            print(f"=== 차트 위치 조정 결과 ===")
            print(f"차트: {chart_name_str}")
            print(f"시트: {sheet_name_str}")
            print()

            print("📍 위치 변경:")
            print(f"   이전: ({original_position['left']:.1f}, {original_position['top']:.1f})")
            print(f"   현재: ({current_pos['left']:.1f}, {current_pos['top']:.1f})")

            print("📏 크기 변경:")
            print(f"   이전: {original_position['width']:.1f} x {original_position['height']:.1f}")
            print(f"   현재: {current_pos['width']:.1f} x {current_pos['height']:.1f}")
            print()

            if changes_made:
                print("✅ 적용된 변경사항:")
                for prop, value in changes_made.items():
                    print(f"   {prop}: {value:.1f}")
                print()

            method = response_data["positioning_method"]
            print(f"🎯 설정 방법: {method['type']}")
            if method["type"] == "relative":
                print(f"   기준 도형: {method['relative_to']}")
                print(f"   방향: {method['direction']}")
                print(f"   간격: {method['offset']}px")
            elif method["type"] == "cell_anchor":
                print(f"   기준 셀: {method['anchor_cell']}")

            if save and file_path:
                print("\n💾 파일이 저장되었습니다.")

    except Exception as e:
        error_response = create_error_response(e, "chart-position")
        fmt = str(output_format) if output_format else "json"
        if fmt == "json":
            print(json.dumps(error_response, ensure_ascii=False, indent=2))
        else:
            print(f"오류: {str(e)}")
        return 1

    finally:
        # COM resource cleanup
        try:
            import gc

            gc.collect()
        except:
            pass

    return 0


if __name__ == "__main__":
    typer.run(chart_position)
