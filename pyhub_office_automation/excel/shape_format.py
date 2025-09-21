"""
도형 스타일 설정 명령어
xlwings를 활용한 Excel 도형 포맷팅 기능
뉴모피즘 스타일 및 고급 그래픽 효과 지원
"""

import json
import platform
from typing import Optional

import typer
import xlwings as xw

from pyhub_office_automation.version import get_version

from .utils import (
    NEUMORPHISM_STYLES,
    ExecutionTimer,
    apply_neumorphism_style,
    create_error_response,
    create_success_response,
    get_or_open_workbook,
    get_shape_by_name,
    get_sheet,
    hex_to_rgb,
    normalize_path,
)


def shape_format(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="도형을 포맷팅할 Excel 파일의 절대 경로"),
    use_active: bool = typer.Option(False, "--use-active", help="현재 활성 워크북 사용"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help='열린 워크북 이름으로 접근 (예: "Sales.xlsx")'),
    sheet: Optional[str] = typer.Option(None, "--sheet", help="도형이 있는 시트 이름 (지정하지 않으면 활성 시트)"),
    shape_name: str = typer.Option(..., "--shape-name", help="포맷팅할 도형 이름"),
    style_preset: str = typer.Option(
        "none", "--style-preset", help="뉴모피즘 스타일 프리셋 적용 (none, background, title-box, chart-box, slicer-box)"
    ),
    fill_color: Optional[str] = typer.Option(None, "--fill-color", help="채우기 색상 (HEX 형식, 예: #FFFFFF)"),
    transparency: Optional[int] = typer.Option(None, "--transparency", help="투명도 (0-100, 0=불투명)"),
    line_color: Optional[str] = typer.Option(None, "--line-color", help="테두리 색상 (HEX 형식)"),
    line_width: Optional[float] = typer.Option(None, "--line-width", help="테두리 두께 (포인트 단위)"),
    no_line: bool = typer.Option(False, "--no-line", help="테두리 제거"),
    shadow_type: str = typer.Option("none", "--shadow-type", help="그림자 타입 (none, drop, inner, outer)"),
    shadow_color: Optional[str] = typer.Option(None, "--shadow-color", help="그림자 색상 (HEX 형식)"),
    shadow_transparency: Optional[int] = typer.Option(None, "--shadow-transparency", help="그림자 투명도 (0-100)"),
    shadow_blur: Optional[int] = typer.Option(None, "--shadow-blur", help="그림자 흐림 정도 (포인트)"),
    shadow_distance: Optional[int] = typer.Option(None, "--shadow-distance", help="그림자 거리 (포인트)"),
    shadow_angle: Optional[int] = typer.Option(None, "--shadow-angle", help="그림자 각도 (도, 0-360)"),
    gradient: bool = typer.Option(False, "--gradient", help="그라데이션 적용 (Windows 전용)"),
    gradient_color2: Optional[str] = typer.Option(None, "--gradient-color2", help="그라데이션 두 번째 색상 (HEX 형식)"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택 (json/text)"),
    visible: bool = typer.Option(False, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부 (기본값: False)"),
    save: bool = typer.Option(True, "--save", help="포맷팅 후 파일 저장 여부 (기본값: True)"),
):
    """
    Excel 도형의 스타일과 포맷을 설정합니다.

    색상, 투명도, 테두리, 그림자 등 다양한 시각적 속성을 조정할 수 있으며,
    뉴모피즘 디자인 시스템에 최적화된 프리셋을 제공합니다.

    === 워크북 접근 방법 ===
    - --file-path: 파일 경로로 워크북 열기
    - --use-active: 현재 활성 워크북 사용
    - --workbook-name: 열린 워크북 이름으로 접근 (예: "Sales.xlsx")

    === 기본 스타일링 ===
    • --fill-color: 채우기 색상 (예: #FFFFFF, #1D2433)
    • --transparency: 투명도 0(불투명) - 100(완전투명)
    • --line-color: 테두리 색상
    • --line-width: 테두리 두께 (포인트)
    • --no-line: 테두리 완전 제거

    === 뉴모피즘 스타일 프리셋 ===
    • background: 배경색 (#F2EDF3)
    • title-box: 제목 영역 (#1D2433, 그림자)
    • chart-box: 차트 영역 (#FFFFFF, 그림자)
    • slicer-box: 슬라이서 영역 (#FFFFFF, 테두리, 그림자)

    === 그림자 효과 (Windows 전용) ===
    • --shadow-type: none, drop, inner, outer
    • --shadow-color: 그림자 색상
    • --shadow-transparency: 그림자 투명도
    • --shadow-blur: 흐림 효과 강도
    • --shadow-distance: 그림자 거리
    • --shadow-angle: 그림자 방향 (각도)

    === 고급 효과 (Windows 전용) ===
    • --gradient: 그라데이션 효과
    • --gradient-color2: 그라데이션 두 번째 색상

    === 실제 사용 시나리오 ===

    # 1. 뉴모피즘 제목 박스 스타일 적용
    oa excel shape-format --use-active --shape-name "TitleArea" --style-preset "title-box"

    # 2. 커스텀 색상으로 차트 영역 스타일링
    oa excel shape-format --use-active --shape-name "ChartArea" \\
        --fill-color "#F8F9FA" --line-color "#DEE2E6" --line-width 1.5

    # 3. 고급 그림자 효과 적용
    oa excel shape-format --use-active --shape-name "HighlightBox" \\
        --fill-color "#FFFFFF" --shadow-type "drop" --shadow-color "#1D2433" \\
        --shadow-transparency 80 --shadow-blur 25 --shadow-distance 8 --shadow-angle 135

    # 4. 투명 오버레이 효과
    oa excel shape-format --use-active --shape-name "Overlay" \\
        --fill-color "#000000" --transparency 50 --no-line

    # 5. 그라데이션 배경 (Windows)
    oa excel shape-format --use-active --shape-name "Background" \\
        --fill-color "#E3F2FD" --gradient --gradient-color2 "#BBDEFB"

    === 대시보드 일관성 관리 ===

    # 모든 차트 영역을 동일한 스타일로 통일
    oa excel shape-format --use-active --shape-name "Chart1Area" --style-preset "chart-box"
    oa excel shape-format --use-active --shape-name "Chart2Area" --style-preset "chart-box"
    oa excel shape-format --use-active --shape-name "Chart3Area" --style-preset "chart-box"

    # 제목과 부제목 영역 차별화
    oa excel shape-format --use-active --shape-name "MainTitle" --style-preset "title-box"
    oa excel shape-format --use-active --shape-name "SubTitle" \\
        --fill-color "#6C757D" --transparency 10

    === 플랫폼별 지원 ===
    • Windows: 모든 기능 지원 (그림자, 그라데이션 포함)
    • macOS: 기본 색상 및 투명도만 지원

    === 색상 가이드 ===
    • 뉴모피즘 배경: #F2EDF3
    • 다크 제목: #1D2433
    • 밝은 컨텐츠: #FFFFFF
    • 경계선: #E0E0E0
    • 강조색: #FF5733, #28A745, #007BFF
    """
    book = None

    try:
        with ExecutionTimer() as timer:
            # 워크북 연결
            book = get_or_open_workbook(
                file_path=file_path, workbook_name=workbook_name, use_active=use_active, visible=visible
            )

            # 시트 가져오기
            target_sheet = get_sheet(book, sheet)

            # 도형 찾기
            shape_obj = get_shape_by_name(target_sheet, shape_name)
            if not shape_obj:
                raise ValueError(f"도형 '{shape_name}'을 찾을 수 없습니다")

            # 적용된 스타일 기록
            applied_styles = []

            # 뉴모피즘 프리셋 적용
            if style_preset and style_preset != "none":
                if style_preset in NEUMORPHISM_STYLES:
                    success = apply_neumorphism_style(shape_obj, style_preset)
                    if success:
                        applied_styles.append(f"뉴모피즘 프리셋 '{style_preset}' 적용")
                    else:
                        applied_styles.append(f"뉴모피즘 프리셋 '{style_preset}' 적용 실패")

            # Windows에서만 고급 포맷팅 지원
            if platform.system() == "Windows":
                try:
                    # 채우기 색상
                    if fill_color:
                        shape_obj.api.Fill.ForeColor.RGB = hex_to_rgb(fill_color)
                        applied_styles.append(f"채우기 색상: {fill_color}")

                    # 투명도
                    if transparency is not None:
                        if 0 <= transparency <= 100:
                            shape_obj.api.Fill.Transparency = transparency / 100.0
                            applied_styles.append(f"투명도: {transparency}%")
                        else:
                            raise ValueError("투명도는 0-100 범위여야 합니다")

                    # 테두리 제거
                    if no_line:
                        shape_obj.api.Line.Visible = False
                        applied_styles.append("테두리 제거")
                    else:
                        # 테두리 색상
                        if line_color:
                            shape_obj.api.Line.Visible = True
                            shape_obj.api.Line.ForeColor.RGB = hex_to_rgb(line_color)
                            applied_styles.append(f"테두리 색상: {line_color}")

                        # 테두리 두께
                        if line_width is not None:
                            shape_obj.api.Line.Visible = True
                            shape_obj.api.Line.Weight = line_width
                            applied_styles.append(f"테두리 두께: {line_width}pt")

                    # 그림자 효과
                    if shadow_type:
                        if shadow_type == "none":
                            shape_obj.api.Shadow.Type = 0  # msoShadowNone
                            applied_styles.append("그림자 제거")
                        else:
                            # 그림자 타입 매핑
                            shadow_type_map = {"drop": 25, "inner": 4, "outer": 25}  # msoShadow25  # msoShadow4  # msoShadow25
                            shape_obj.api.Shadow.Type = shadow_type_map.get(shadow_type, 25)
                            applied_styles.append(f"그림자 타입: {shadow_type}")

                            # 그림자 상세 설정
                            if shadow_color:
                                shape_obj.api.Shadow.ForeColor.RGB = hex_to_rgb(shadow_color)
                                applied_styles.append(f"그림자 색상: {shadow_color}")

                            if shadow_transparency is not None:
                                if 0 <= shadow_transparency <= 100:
                                    shape_obj.api.Shadow.Transparency = shadow_transparency / 100.0
                                    applied_styles.append(f"그림자 투명도: {shadow_transparency}%")

                            if shadow_blur is not None:
                                shape_obj.api.Shadow.Blur = shadow_blur
                                applied_styles.append(f"그림자 흐림: {shadow_blur}pt")

                            if shadow_distance is not None:
                                shape_obj.api.Shadow.OffsetX = shadow_distance
                                shape_obj.api.Shadow.OffsetY = shadow_distance
                                applied_styles.append(f"그림자 거리: {shadow_distance}pt")

                            if shadow_angle is not None:
                                # 각도를 라디안으로 변환하여 X, Y 오프셋 계산
                                import math

                                angle_rad = math.radians(shadow_angle)
                                distance = shadow_distance or 5
                                shape_obj.api.Shadow.OffsetX = distance * math.cos(angle_rad)
                                shape_obj.api.Shadow.OffsetY = distance * math.sin(angle_rad)
                                applied_styles.append(f"그림자 각도: {shadow_angle}도")

                    # 그라데이션 효과
                    if gradient:
                        try:
                            # 그라데이션 타입 설정
                            shape_obj.api.Fill.TwoColorGradient(1, 1)  # msoGradientHorizontal, msoGradientVariant1

                            if gradient_color2:
                                shape_obj.api.Fill.BackColor.RGB = hex_to_rgb(gradient_color2)
                                applied_styles.append(f"그라데이션 적용: {fill_color or '기본색'} → {gradient_color2}")
                            else:
                                applied_styles.append("그라데이션 적용")
                        except Exception:
                            applied_styles.append("그라데이션 적용 실패")

                except Exception as e:
                    raise RuntimeError(f"스타일 적용 중 오류 발생: {str(e)}")

            else:
                # macOS에서는 제한된 기능만 지원
                if fill_color or transparency is not None:
                    applied_styles.append("macOS에서는 제한된 스타일링만 지원됩니다")

            # 파일 저장
            if save and file_path:
                book.save()

            # 현재 도형 정보 수집
            current_info = {
                "name": shape_obj.name,
                "position": {"left": getattr(shape_obj, "left", 0), "top": getattr(shape_obj, "top", 0)},
                "size": {"width": getattr(shape_obj, "width", 0), "height": getattr(shape_obj, "height", 0)},
            }

            # Windows에서 현재 스타일 정보 수집
            if platform.system() == "Windows":
                try:
                    from .utils import rgb_to_hex

                    current_info["current_style"] = {
                        "fill_color": rgb_to_hex(shape_obj.api.Fill.ForeColor.RGB),
                        "transparency": round(shape_obj.api.Fill.Transparency * 100),
                        "has_line": shape_obj.api.Line.Visible,
                        "has_shadow": shape_obj.api.Shadow.Type != 0,
                    }

                    if current_info["current_style"]["has_line"]:
                        current_info["current_style"]["line_color"] = rgb_to_hex(shape_obj.api.Line.ForeColor.RGB)

                    if current_info["current_style"]["has_shadow"]:
                        current_info["current_style"]["shadow_color"] = rgb_to_hex(shape_obj.api.Shadow.ForeColor.RGB)

                except Exception:
                    pass

            # 성공 응답 생성
            response_data = {
                "shape_info": current_info,
                "applied_styles": applied_styles,
                "total_styles_applied": len([s for s in applied_styles if "실패" not in s]),
                "sheet": target_sheet.name,
                "workbook": normalize_path(book.name),
                "platform_support": "full" if platform.system() == "Windows" else "limited",
            }

            # 적용된 스타일 요약
            style_summary = {}
            if style_preset and style_preset != "none":
                style_summary["preset"] = style_preset
            if fill_color:
                style_summary["fill_color"] = fill_color
            if transparency is not None:
                style_summary["transparency"] = transparency
            if shadow_type:
                style_summary["shadow_type"] = shadow_type

            if style_summary:
                response_data["style_summary"] = style_summary

            message = (
                f"도형 '{shape_name}'에 {len([s for s in applied_styles if '실패' not in s])}가지 스타일이 적용되었습니다"
            )

            response = create_success_response(
                data=response_data,
                command="shape-format",
                message=message,
                execution_time_ms=timer.execution_time_ms,
                book=book,
                styles_applied=len(applied_styles),
            )

            print(json.dumps(response, ensure_ascii=False, indent=2))

    except Exception as e:
        error_response = create_error_response(e, "shape-format")
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
    shape_format()
