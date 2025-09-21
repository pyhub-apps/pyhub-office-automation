"""
슬라이서 위치 조정 명령어
xlwings를 활용한 Excel 슬라이서 위치 및 크기 조정 기능
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
    get_slicer_by_name,
    normalize_path,
    validate_slicer_position,
)


def slicer_position(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="슬라이서 위치를 조정할 Excel 파일의 절대 경로"),
    use_active: bool = typer.Option(False, "--use-active", help="현재 활성 워크북 사용"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help='열린 워크북 이름으로 접근 (예: "Sales.xlsx")'),
    slicer_name: str = typer.Option(..., "--slicer-name", help="위치를 조정할 슬라이서 이름"),
    left: Optional[int] = typer.Option(None, "--left", help="새로운 왼쪽 위치 (픽셀)"),
    top: Optional[int] = typer.Option(None, "--top", help="새로운 위쪽 위치 (픽셀)"),
    width: Optional[int] = typer.Option(None, "--width", help="새로운 너비 (픽셀)"),
    height: Optional[int] = typer.Option(None, "--height", help="새로운 높이 (픽셀)"),
    move_by_x: Optional[int] = typer.Option(None, "--move-by-x", help="X축 상대 이동 거리 (픽셀, 양수=오른쪽, 음수=왼쪽)"),
    move_by_y: Optional[int] = typer.Option(None, "--move-by-y", help="Y축 상대 이동 거리 (픽셀, 양수=아래, 음수=위)"),
    resize_by_width: Optional[int] = typer.Option(
        None, "--resize-by-width", help="너비 상대 조정 (픽셀, 양수=확대, 음수=축소)"
    ),
    resize_by_height: Optional[int] = typer.Option(
        None, "--resize-by-height", help="높이 상대 조정 (픽셀, 양수=확대, 음수=축소)"
    ),
    align_to: Optional[str] = typer.Option(
        None, "--align-to", help="정렬 기준 (left/center/right/top/middle/bottom, 워크시트 기준)"
    ),
    align_with: Optional[str] = typer.Option(None, "--align-with", help="다른 슬라이서와 정렬 (슬라이서 이름)"),
    preset_layout: Optional[str] = typer.Option(
        None, "--preset-layout", help="미리 정의된 레이아웃 적용 (horizontal/vertical/grid-2x2/grid-3x1/sidebar)"
    ),
    snap_to_grid: Optional[int] = typer.Option(None, "--snap-to-grid", help="그리드에 맞춤 (픽셀 단위, 예: 10)"),
    auto_arrange: bool = typer.Option(False, "--auto-arrange", help="다른 슬라이서와 겹치지 않도록 자동 배치"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택 (json/text)"),
    visible: bool = typer.Option(False, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부 (기본값: False)"),
    save: bool = typer.Option(True, "--save", help="위치 조정 후 파일 저장 여부 (기본값: True)"),
):
    """
    Excel 슬라이서의 위치와 크기를 조정합니다.

    절대 위치 지정, 상대 이동, 정렬, 미리 정의된 레이아웃 등 다양한 방법으로
    슬라이서를 정확하게 배치할 수 있으며, 대시보드 레이아웃 최적화에 유용합니다.

    === 워크북 접근 방법 ===
    - --file-path: 파일 경로로 워크북 열기
    - --use-active: 현재 활성 워크북 사용
    - --workbook-name: 열린 워크북 이름으로 접근 (예: "Sales.xlsx")

    === 절대 위치 지정 ===
    • --left, --top: 새로운 절대 위치 (픽셀)
    • --width, --height: 새로운 절대 크기 (픽셀)

    === 상대 이동 ===
    • --move-by-x: X축 상대 이동 (양수=오른쪽, 음수=왼쪽)
    • --move-by-y: Y축 상대 이동 (양수=아래, 음수=위)
    • --resize-by-width: 너비 상대 조정
    • --resize-by-height: 높이 상대 조정

    === 정렬 및 배치 ===
    • --align-to: 워크시트 기준 정렬
    • --align-with: 다른 슬라이서와 정렬
    • --preset-layout: 미리 정의된 레이아웃
    • --snap-to-grid: 그리드에 맞춤
    • --auto-arrange: 자동 배치 (겹침 방지)

    === 절대 위치 지정 시나리오 ===

    # 1. 슬라이서를 특정 위치로 정확히 이동
    oa excel slicer-position --use-active --slicer-name "RegionSlicer" \\
        --left 100 --top 400 --width 200 --height 150

    # 2. 위치만 변경 (크기 유지)
    oa excel slicer-position --use-active --slicer-name "MonthSlicer" \\
        --left 320 --top 400

    # 3. 크기만 조정 (위치 유지)
    oa excel slicer-position --use-active --slicer-name "CategorySlicer" \\
        --width 300 --height 100

    === 상대 이동 시나리오 ===

    # 1. 슬라이서를 오른쪽으로 50픽셀 이동
    oa excel slicer-position --use-active --slicer-name "ProductSlicer" \\
        --move-by-x 50

    # 2. 슬라이서를 위로 30픽셀, 왼쪽으로 20픽셀 이동
    oa excel slicer-position --use-active --slicer-name "DateSlicer" \\
        --move-by-x -20 --move-by-y -30

    # 3. 슬라이서 크기를 가로 50픽셀, 세로 20픽셀 확대
    oa excel slicer-position --use-active --slicer-name "SalesPersonSlicer" \\
        --resize-by-width 50 --resize-by-height 20

    # 4. 슬라이서를 왼쪽으로 이동하면서 크기 축소
    oa excel slicer-position --use-active --slicer-name "TempSlicer" \\
        --move-by-x -100 --resize-by-width -50

    === 정렬 시나리오 ===

    # 1. 워크시트 왼쪽에 정렬
    oa excel slicer-position --use-active --slicer-name "RegionSlicer" \\
        --align-to left

    # 2. 워크시트 중앙에 정렬
    oa excel slicer-position --use-active --slicer-name "MainSlicer" \\
        --align-to center

    # 3. 다른 슬라이서와 정렬 (왼쪽 맞춤)
    oa excel slicer-position --use-active --slicer-name "SubSlicer" \\
        --align-with "MainSlicer"

    === 미리 정의된 레이아웃 ===

    # 1. 수평 배치 레이아웃 (가로로 나란히)
    oa excel slicer-position --use-active --slicer-name "RegionSlicer" \\
        --preset-layout horizontal

    # 2. 수직 배치 레이아웃 (세로로 나란히)
    oa excel slicer-position --use-active --slicer-name "CategorySlicer" \\
        --preset-layout vertical

    # 3. 2x2 그리드 레이아웃
    oa excel slicer-position --use-active --slicer-name "DateSlicer" \\
        --preset-layout grid-2x2

    # 4. 사이드바 배치
    oa excel slicer-position --use-active --slicer-name "FilterSlicer" \\
        --preset-layout sidebar

    === 고급 배치 및 정렬 ===

    # 1. 그리드에 맞춤 (10픽셀 단위)
    oa excel slicer-position --use-active --slicer-name "RegionSlicer" \\
        --left 105 --top 407 --snap-to-grid 10

    # 2. 자동 배치 (다른 슬라이서와 겹치지 않음)
    oa excel slicer-position --use-active --slicer-name "NewSlicer" \\
        --auto-arrange --width 200 --height 120

    # 3. 복합 조정 (이동 + 크기 조정 + 그리드 맞춤)
    oa excel slicer-position --use-active --slicer-name "MainSlicer" \\
        --move-by-x 25 --resize-by-width 30 --snap-to-grid 5

    === 대시보드 레이아웃 구성 ===

    # 뉴모피즘 슬라이서 영역 내부 정렬
    # 1. 첫 번째 슬라이서 - 왼쪽 정렬
    oa excel slicer-position --use-active --slicer-name "RegionSlicer" \\
        --left 100 --top 420 --width 150 --height 80

    # 2. 두 번째 슬라이서 - 첫 번째 오른쪽에 배치
    oa excel slicer-position --use-active --slicer-name "MonthSlicer" \\
        --left 270 --top 420 --width 150 --height 80

    # 3. 세 번째 슬라이서 - 두 번째 오른쪽에 배치
    oa excel slicer-position --use-active --slicer-name "CategorySlicer" \\
        --left 440 --top 420 --width 150 --height 80

    # 4. 네 번째 슬라이서 - 세 번째 오른쪽에 배치
    oa excel slicer-position --use-active --slicer-name "ProductSlicer" \\
        --left 610 --top 420 --width 150 --height 80

    === 반응형 레이아웃 조정 ===

    # 화면 크기에 따른 슬라이서 재배치
    # 1. 큰 화면용 - 가로 배치
    for slicer in ["Region", "Month", "Category", "Product"]:
        oa excel slicer-position --use-active --slicer-name f"{slicer}Slicer" \\
            --preset-layout horizontal

    # 2. 작은 화면용 - 세로 배치
    for slicer in ["Region", "Month", "Category", "Product"]:
        oa excel slicer-position --use-active --slicer-name f"{slicer}Slicer" \\
            --preset-layout vertical

    === 배치 최적화 팁 ===
    • 슬라이서 간 최소 10-20픽셀 간격 유지
    • 그리드 시스템 사용으로 정돈된 레이아웃 구성
    • 사용 빈도에 따른 우선순위 배치
    • 화면 해상도 고려한 반응형 설계
    • 논리적 그룹핑으로 사용자 경험 향상

    === 주의사항 ===
    • Windows에서만 완전 지원
    • 슬라이서 겹침 시 상위 레이어가 하위 레이어 가림
    • 너무 작은 크기는 사용성 저하 (최소 100x50픽셀 권장)
    • 워크시트 경계를 벗어나지 않도록 주의
    • 다른 차트나 도형과의 조화 고려
    """
    book = None

    try:
        # 매개변수 검증
        if align_to and align_to not in ["left", "center", "right", "top", "middle", "bottom"]:
            raise ValueError(
                f"align_to는 'left', 'center', 'right', 'top', 'middle', 'bottom' 중 하나여야 합니다. 입력된 값: {align_to}"
            )

        if preset_layout and preset_layout not in ["horizontal", "vertical", "grid-2x2", "grid-3x1", "sidebar"]:
            raise ValueError(
                f"preset_layout은 'horizontal', 'vertical', 'grid-2x2', 'grid-3x1', 'sidebar' 중 하나여야 합니다. 입력된 값: {preset_layout}"
            )

        with ExecutionTimer() as timer:
            # Windows 플랫폼 확인
            if platform.system() != "Windows":
                raise RuntimeError("슬라이서 위치 조정은 Windows에서만 지원됩니다")

            # 워크북 연결
            book = get_or_open_workbook(
                file_path=file_path, workbook_name=workbook_name, use_active=use_active, visible=visible
            )

            # 슬라이서 찾기
            slicer_cache = get_slicer_by_name(book, slicer_name)
            if not slicer_cache:
                raise ValueError(f"슬라이서 '{slicer_name}'을 찾을 수 없습니다")

            # 첫 번째 슬라이서 객체 가져오기 (위치 조정 대상)
            if slicer_cache.Slicers().Count == 0:
                raise ValueError(f"슬라이서 '{slicer_name}'에 시각적 슬라이서가 없습니다")

            slicer_obj = slicer_cache.Slicers(1)

            # 현재 위치 및 크기 정보
            current_position = {
                "left": slicer_obj.Left,
                "top": slicer_obj.Top,
                "width": slicer_obj.Width,
                "height": slicer_obj.Height,
            }

            # 새로운 위치 및 크기 계산
            new_position = calculate_new_position(
                current_position,
                left,
                top,
                width,
                height,
                move_by_x,
                move_by_y,
                resize_by_width,
                resize_by_height,
                align_to,
                align_with,
                preset_layout,
                snap_to_grid,
                auto_arrange,
                slicer_obj,
                book,
            )

            # 위치 및 크기 유효성 검증
            is_valid, error_msg = validate_slicer_position(
                new_position["left"], new_position["top"], new_position["width"], new_position["height"]
            )
            if not is_valid:
                raise ValueError(error_msg)

            # 실제 위치 조정 수행
            changes_made = []

            if new_position["left"] != current_position["left"]:
                slicer_obj.Left = new_position["left"]
                changes_made.append(f"left: {current_position['left']} → {new_position['left']}")

            if new_position["top"] != current_position["top"]:
                slicer_obj.Top = new_position["top"]
                changes_made.append(f"top: {current_position['top']} → {new_position['top']}")

            if new_position["width"] != current_position["width"]:
                slicer_obj.Width = new_position["width"]
                changes_made.append(f"width: {current_position['width']} → {new_position['width']}")

            if new_position["height"] != current_position["height"]:
                slicer_obj.Height = new_position["height"]
                changes_made.append(f"height: {current_position['height']} → {new_position['height']}")

            if not changes_made:
                changes_made.append("변경사항 없음 (이미 목표 위치/크기)")

            # 파일 저장
            if save and file_path:
                book.save()

            # 성공 응답 생성
            response_data = {
                "slicer_name": slicer_name,
                "previous_position": current_position,
                "new_position": new_position,
                "changes_made": changes_made,
                "total_changes": len([c for c in changes_made if "→" in c]),
                "workbook": normalize_path(book.name),
                "sheet": slicer_obj.Parent.Parent.Name,
            }

            # 조정 방법 정보
            adjustment_methods = []
            if left is not None or top is not None or width is not None or height is not None:
                adjustment_methods.append("절대 위치 지정")
            if move_by_x or move_by_y or resize_by_width or resize_by_height:
                adjustment_methods.append("상대 조정")
            if align_to:
                adjustment_methods.append(f"정렬 ({align_to})")
            if align_with:
                adjustment_methods.append(f"다른 슬라이서와 정렬 ({align_with})")
            if preset_layout:
                adjustment_methods.append(f"프리셋 레이아웃 ({preset_layout})")
            if snap_to_grid:
                adjustment_methods.append(f"그리드 맞춤 ({snap_to_grid}px)")
            if auto_arrange:
                adjustment_methods.append("자동 배치")

            if adjustment_methods:
                response_data["adjustment_methods"] = adjustment_methods

            message = f"슬라이서 '{slicer_name}'의 위치가 성공적으로 조정되었습니다"
            if len([c for c in changes_made if "→" in c]) == 0:
                message = f"슬라이서 '{slicer_name}'는 이미 목표 위치에 있습니다"

            response = create_success_response(
                data=response_data,
                command="slicer-position",
                message=message,
                execution_time_ms=timer.execution_time_ms,
                book=book,
                changes_count=len([c for c in changes_made if "→" in c]),
            )

            print(json.dumps(response, ensure_ascii=False, indent=2))

    except Exception as e:
        error_response = create_error_response(e, "slicer-position")
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


def calculate_new_position(
    current_position,
    left,
    top,
    width,
    height,
    move_by_x,
    move_by_y,
    resize_by_width,
    resize_by_height,
    align_to,
    align_with,
    preset_layout,
    snap_to_grid,
    auto_arrange,
    slicer_obj,
    book,
):
    """새로운 위치와 크기를 계산합니다."""

    new_position = current_position.copy()

    # 1. 절대 위치 지정
    if left is not None:
        new_position["left"] = left
    if top is not None:
        new_position["top"] = top
    if width is not None:
        new_position["width"] = width
    if height is not None:
        new_position["height"] = height

    # 2. 상대 이동
    if move_by_x:
        new_position["left"] += move_by_x
    if move_by_y:
        new_position["top"] += move_by_y
    if resize_by_width:
        new_position["width"] += resize_by_width
    if resize_by_height:
        new_position["height"] += resize_by_height

    # 3. 정렬 처리
    if align_to:
        sheet = slicer_obj.Parent.Parent
        sheet_width = sheet.Cells.SpecialCells(11).Column * 64  # 대략적인 시트 너비
        sheet_height = sheet.Cells.SpecialCells(11).Row * 20  # 대략적인 시트 높이

        if align_to == "left":
            new_position["left"] = 50  # 왼쪽 여백
        elif align_to == "center":
            new_position["left"] = (sheet_width - new_position["width"]) // 2
        elif align_to == "right":
            new_position["left"] = sheet_width - new_position["width"] - 50
        elif align_to == "top":
            new_position["top"] = 50  # 위쪽 여백
        elif align_to == "middle":
            new_position["top"] = (sheet_height - new_position["height"]) // 2
        elif align_to == "bottom":
            new_position["top"] = sheet_height - new_position["height"] - 50

    # 4. 다른 슬라이서와 정렬
    if align_with:
        target_slicer_cache = get_slicer_by_name(book, align_with)
        if target_slicer_cache and target_slicer_cache.Slicers().Count > 0:
            target_slicer = target_slicer_cache.Slicers(1)
            new_position["left"] = target_slicer.Left
            new_position["top"] = target_slicer.Top

    # 5. 프리셋 레이아웃
    if preset_layout:
        new_position = apply_preset_layout(new_position, preset_layout, slicer_obj)

    # 6. 그리드 맞춤
    if snap_to_grid:
        new_position["left"] = round(new_position["left"] / snap_to_grid) * snap_to_grid
        new_position["top"] = round(new_position["top"] / snap_to_grid) * snap_to_grid

    # 7. 자동 배치 (겹침 방지)
    if auto_arrange:
        new_position = avoid_overlap(new_position, slicer_obj, book)

    return new_position


def apply_preset_layout(position, layout_type, slicer_obj):
    """미리 정의된 레이아웃을 적용합니다."""

    if layout_type == "horizontal":
        # 수평 배치 - 워크시트 상단에 가로로 배열
        position["left"] = 100
        position["top"] = 50
        position["width"] = 180
        position["height"] = 100

    elif layout_type == "vertical":
        # 수직 배치 - 워크시트 왼쪽에 세로로 배열
        position["left"] = 50
        position["top"] = 100
        position["width"] = 150
        position["height"] = 120

    elif layout_type == "grid-2x2":
        # 2x2 그리드 - 첫 번째 위치
        position["left"] = 100
        position["top"] = 100
        position["width"] = 160
        position["height"] = 100

    elif layout_type == "grid-3x1":
        # 3x1 그리드 - 첫 번째 위치
        position["left"] = 100
        position["top"] = 400
        position["width"] = 200
        position["height"] = 80

    elif layout_type == "sidebar":
        # 사이드바 배치 - 오른쪽
        position["left"] = 800
        position["top"] = 100
        position["width"] = 180
        position["height"] = 500

    return position


def avoid_overlap(position, slicer_obj, book):
    """다른 슬라이서와 겹치지 않도록 위치를 조정합니다."""

    try:
        # 다른 슬라이서들의 위치 정보 수집
        other_slicers = []
        for slicer_cache in book.api.SlicerCaches():
            if slicer_cache.Name != slicer_obj.Parent.Name:
                if slicer_cache.Slicers().Count > 0:
                    other_slicer = slicer_cache.Slicers(1)
                    other_slicers.append(
                        {
                            "left": other_slicer.Left,
                            "top": other_slicer.Top,
                            "width": other_slicer.Width,
                            "height": other_slicer.Height,
                        }
                    )

        # 겹침 검사 및 위치 조정
        adjusted = False
        max_attempts = 10
        attempts = 0

        while attempts < max_attempts:
            overlapping = False

            for other in other_slicers:
                if is_overlapping(position, other):
                    # 겹치는 경우 오른쪽으로 이동
                    position["left"] = other["left"] + other["width"] + 20
                    overlapping = True
                    adjusted = True
                    break

            if not overlapping:
                break

            attempts += 1

        # 화면 경계 확인
        if position["left"] > 1000:  # 너무 오른쪽으로 밀린 경우
            position["left"] = 100
            position["top"] += position["height"] + 20  # 아래로 이동

    except Exception:
        # 자동 배치 실패 시 기본 위치 유지
        pass

    return position


def is_overlapping(pos1, pos2):
    """두 영역이 겹치는지 확인합니다."""
    return not (
        pos1["left"] + pos1["width"] <= pos2["left"]
        or pos2["left"] + pos2["width"] <= pos1["left"]
        or pos1["top"] + pos1["height"] <= pos2["top"]
        or pos2["top"] + pos2["height"] <= pos1["top"]
    )


if __name__ == "__main__":
    slicer_position()
