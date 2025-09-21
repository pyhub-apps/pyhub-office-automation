"""
도형 추가 명령어
xlwings를 활용한 Excel 도형 생성 기능
뉴모피즘 스타일 대시보드 지원
"""

import json
import platform
from typing import Optional

import typer
import xlwings as xw

from pyhub_office_automation.version import get_version

from .utils import (
    NEUMORPHISM_STYLES,
    SHAPE_TYPES,
    ExecutionTimer,
    apply_neumorphism_style,
    create_error_response,
    create_success_response,
    generate_unique_shape_name,
    get_or_open_workbook,
    get_sheet,
    normalize_path,
    validate_position_and_size,
)


def shape_add(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="도형을 추가할 Excel 파일의 절대 경로"),
    use_active: bool = typer.Option(False, "--use-active", help="현재 활성 워크북 사용"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help='열린 워크북 이름으로 접근 (예: "Sales.xlsx")'),
    sheet: Optional[str] = typer.Option(None, "--sheet", help="도형을 추가할 시트 이름 (지정하지 않으면 활성 시트)"),
    shape_type: str = typer.Option(
        "rectangle", "--shape-type", help="도형 유형 (기본값: rectangle) - rectangle, oval, line, arrow, textbox 등 사용 가능"
    ),
    left: int = typer.Option(100, "--left", help="도형의 왼쪽 위치 (픽셀, 기본값: 100)"),
    top: int = typer.Option(100, "--top", help="도형의 위쪽 위치 (픽셀, 기본값: 100)"),
    width: int = typer.Option(200, "--width", help="도형의 너비 (픽셀, 기본값: 200)"),
    height: int = typer.Option(100, "--height", help="도형의 높이 (픽셀, 기본값: 100)"),
    name: Optional[str] = typer.Option(None, "--name", help="도형 이름 (지정하지 않으면 자동 생성)"),
    style_preset: str = typer.Option(
        "none",
        "--style-preset",
        help="뉴모피즘 스타일 프리셋 (기본값: none) - none, background, title-box, chart-box, slicer-box 사용 가능",
    ),
    fill_color: Optional[str] = typer.Option(None, "--fill-color", help="채우기 색상 (HEX 형식, 예: #FFFFFF)"),
    transparency: Optional[int] = typer.Option(None, "--transparency", help="투명도 (0-100, 0=불투명)"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택 (json/text)"),
    visible: bool = typer.Option(False, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부 (기본값: False)"),
    save: bool = typer.Option(True, "--save", help="생성 후 파일 저장 여부 (기본값: True)"),
):
    """
    Excel 시트에 도형을 추가합니다.

    다양한 도형 유형을 지원하며, 뉴모피즘 디자인 시스템과 통합된 대시보드 구성이 가능합니다.
    정확한 픽셀 단위 위치 지정과 스타일 프리셋을 제공합니다.

    === 워크북 접근 방법 ===
    - --file-path: 파일 경로로 워크북 열기
    - --use-active: 현재 활성 워크북 사용
    - --workbook-name: 열린 워크북 이름으로 접근 (예: "Sales.xlsx")

    === 지원되는 도형 유형 ===
    • rectangle: 사각형 (기본값)
    • oval: 원/타원
    • rounded_rectangle: 둥근 사각형
    • triangle: 삼각형
    • diamond: 마름모
    • pentagon: 오각형
    • hexagon: 육각형
    • star: 별 모양
    • arrow: 화살표
    • line: 선
    • callout_rectangle: 말풍선 사각형
    • text_box: 텍스트 박스

    === 위치와 크기 지정 ===
    • --left, --top: 도형의 좌상단 위치 (픽셀 단위)
    • --width, --height: 도형의 크기 (픽셀 단위)
    • 유효 범위: 위치 0-20000px, 크기 1-5000px

    === 뉴모피즘 스타일 프리셋 ===
    • none: 기본 도형 (스타일 없음)
    • background: 배경색 (#F2EDF3)
    • title-box: 제목 영역 (#1D2433, 그림자 효과)
    • chart-box: 차트 영역 (#FFFFFF, 그림자 효과)
    • slicer-box: 슬라이서 영역 (#FFFFFF, 테두리, 그림자 효과)

    === 대시보드 구성 시나리오 ===

    # 1. 배경 도형 생성
    oa excel shape-add --use-active --shape-type rectangle --left 50 --top 50 \\
        --width 800 --height 600 --style-preset background --name "DashboardBG"

    # 2. 제목 영역 추가
    oa excel shape-add --use-active --shape-type rounded_rectangle --left 70 --top 70 \\
        --width 760 --height 80 --style-preset title-box --name "TitleArea"

    # 3. 차트 영역들 추가
    oa excel shape-add --use-active --shape-type rounded_rectangle --left 90 --top 170 \\
        --width 350 --height 200 --style-preset chart-box --name "Chart1Area"

    oa excel shape-add --use-active --shape-type rounded_rectangle --left 460 --top 170 \\
        --width 350 --height 200 --style-preset chart-box --name "Chart2Area"

    # 4. 슬라이서 영역 추가
    oa excel shape-add --use-active --shape-type rounded_rectangle --left 90 --top 390 \\
        --width 720 --height 120 --style-preset slicer-box --name "SlicerArea"

    # 5. 커스텀 색상으로 도형 추가
    oa excel shape-add --use-active --shape-type oval --left 200 --top 300 \\
        --width 100 --height 100 --fill-color "#FF5733" --transparency 20

    === 고급 활용 ===
    • 도형 이름 자동 생성: --name을 지정하지 않으면 중복되지 않는 이름 자동 생성
    • 픽셀 정밀도: 정확한 위치 지정으로 그리드 기반 레이아웃 구성
    • 스타일 조합: 프리셋과 개별 색상 옵션 동시 사용 가능
    • 다중 시트: --sheet 옵션으로 특정 시트에 도형 배치
    """
    book = None

    try:
        with ExecutionTimer() as timer:
            # 위치와 크기 검증
            is_valid, error_msg = validate_position_and_size(left, top, width, height)
            if not is_valid:
                raise ValueError(error_msg)

            # 워크북 연결
            book = get_or_open_workbook(
                file_path=file_path, workbook_name=workbook_name, use_active=use_active, visible=visible
            )

            # 시트 가져오기
            target_sheet = get_sheet(book, sheet)

            # 도형 이름 결정
            if name:
                # 중복 확인
                existing_shape = None
                try:
                    for shape in target_sheet.shapes:
                        if shape.name == name:
                            existing_shape = shape
                            break
                except Exception:
                    pass

                if existing_shape:
                    raise ValueError(f"도형 이름 '{name}'이 이미 존재합니다")
                shape_name = name
            else:
                shape_name = generate_unique_shape_name(target_sheet, "Shape")

            # 도형 타입 확인
            if shape_type not in SHAPE_TYPES:
                raise ValueError(f"지원되지 않는 도형 타입: {shape_type}")

            shape_type_value = SHAPE_TYPES[shape_type]

            # 도형 생성
            try:
                if platform.system() == "Windows":
                    # Windows에서는 COM API 사용
                    shape = target_sheet.api.Shapes.AddShape(
                        Type=shape_type_value, Left=left, Top=top, Width=width, Height=height
                    )
                    shape.Name = shape_name
                    # xlwings 객체로 래핑
                    shape_obj = target_sheet.shapes[shape_name]
                else:
                    # macOS에서는 제한적 지원
                    # 경고 메시지와 함께 기본 처리만 제공
                    raise RuntimeError(
                        "macOS에서는 Excel 도형 생성이 제한됩니다. "
                        "이 기능은 Windows 환경에서만 완전히 지원됩니다. "
                        "macOS에서는 수동으로 도형을 추가한 후 다른 명령어(shape-list, shape-format)를 사용하시기 바랍니다."
                    )

            except Exception as e:
                raise RuntimeError(f"도형 생성 실패: {str(e)}")

            # 스타일 적용
            style_applied = False
            if style_preset != "none":
                style_applied = apply_neumorphism_style(shape_obj, style_preset)

            # 개별 색상 설정 (프리셋보다 우선)
            if fill_color:
                try:
                    if platform.system() == "Windows":
                        from .utils import hex_to_rgb

                        shape_obj.api.Fill.ForeColor.RGB = hex_to_rgb(fill_color)
                        style_applied = True
                except Exception:
                    pass

            # 투명도 설정
            if transparency is not None:
                try:
                    if platform.system() == "Windows":
                        shape_obj.api.Fill.Transparency = transparency / 100.0
                        style_applied = True
                except Exception:
                    pass

            # 파일 저장
            if save and file_path:
                book.save()

            # 성공 응답 생성
            response_data = {
                "shape_name": shape_name,
                "shape_type": shape_type,
                "position": {"left": left, "top": top},
                "size": {"width": width, "height": height},
                "sheet": target_sheet.name,
                "workbook": normalize_path(book.name),
            }

            # 스타일 정보 추가
            if style_preset != "none":
                response_data["style_preset"] = style_preset
                response_data["style_applied"] = style_applied

            if fill_color:
                response_data["fill_color"] = fill_color

            if transparency is not None:
                response_data["transparency"] = transparency

            response = create_success_response(
                data=response_data,
                command="shape-add",
                message=f"도형 '{shape_name}'이 성공적으로 생성되었습니다",
                execution_time_ms=timer.execution_time_ms,
                book=book,
                shape_count=len(target_sheet.shapes),
            )

            typer.echo(json.dumps(response, ensure_ascii=False, indent=2))

    except Exception as e:
        error_response = create_error_response(e, "shape-add")
        typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2))
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
