"""
텍스트 박스 추가 명령어
xlwings를 활용한 Excel 텍스트 박스 생성 및 스타일링 기능
"""

import json
import platform
import click
import xlwings as xw
from ..version import get_version
from .utils import (
    get_or_open_workbook, get_sheet, create_error_response, create_success_response,
    ExecutionTimer, validate_position_and_size, generate_unique_shape_name,
    hex_to_rgb, normalize_path
)


@click.command()
@click.option('--file-path',
              help='텍스트 박스를 추가할 Excel 파일의 절대 경로')
@click.option('--use-active', is_flag=True,
              help='현재 활성 워크북 사용')
@click.option('--workbook-name',
              help='열린 워크북 이름으로 접근 (예: "Sales.xlsx")')
@click.option('--sheet',
              help='텍스트 박스를 추가할 시트 이름 (지정하지 않으면 활성 시트)')
@click.option('--text', required=True,
              help='텍스트 박스에 입력할 텍스트 내용')
@click.option('--left', type=int, default=100,
              help='텍스트 박스의 왼쪽 위치 (픽셀, 기본값: 100)')
@click.option('--top', type=int, default=100,
              help='텍스트 박스의 위쪽 위치 (픽셀, 기본값: 100)')
@click.option('--width', type=int, default=200,
              help='텍스트 박스의 너비 (픽셀, 기본값: 200)')
@click.option('--height', type=int, default=50,
              help='텍스트 박스의 높이 (픽셀, 기본값: 50)')
@click.option('--name',
              help='텍스트 박스 이름 (지정하지 않으면 자동 생성)')
@click.option('--font-size', type=int, default=12,
              help='글꼴 크기 (포인트, 기본값: 12)')
@click.option('--font-color', default='#000000',
              help='글꼴 색상 (HEX 형식, 기본값: #000000)')
@click.option('--font-name', default='Arial',
              help='글꼴 이름 (기본값: Arial)')
@click.option('--bold', is_flag=True,
              help='굵은 글꼴')
@click.option('--italic', is_flag=True,
              help='기울임 글꼴')
@click.option('--alignment',
              type=click.Choice(['left', 'center', 'right']),
              default='left',
              help='텍스트 정렬 (기본값: left)')
@click.option('--vertical-alignment',
              type=click.Choice(['top', 'middle', 'bottom']),
              default='top',
              help='수직 정렬 (기본값: top)')
@click.option('--fill-color',
              help='배경 색상 (HEX 형식, 지정하지 않으면 투명)')
@click.option('--transparency', type=int,
              help='배경 투명도 (0-100)')
@click.option('--border-color',
              help='테두리 색상 (HEX 형식)')
@click.option('--border-width', type=float,
              help='테두리 두께 (포인트)')
@click.option('--no-border', is_flag=True,
              help='테두리 제거')
@click.option('--word-wrap', is_flag=True, default=True,
              help='자동 줄바꿈 (기본값: True)')
@click.option('--auto-size', is_flag=True,
              help='내용에 맞게 크기 자동 조정')
@click.option('--format', 'output_format', default='json',
              type=click.Choice(['json', 'text']),
              help='출력 형식 선택')
@click.option('--visible', default=False, type=bool,
              help='Excel 애플리케이션을 화면에 표시할지 여부 (기본값: False)')
@click.option('--save', default=True, type=bool,
              help='생성 후 파일 저장 여부 (기본값: True)')
@click.version_option(version=get_version(), prog_name="oa excel textbox-add")
def textbox_add(file_path, use_active, workbook_name, sheet, text, left, top,
                width, height, name, font_size, font_color, font_name, bold, italic,
                alignment, vertical_alignment, fill_color, transparency, border_color,
                border_width, no_border, word_wrap, auto_size, output_format, visible, save):
    """
    Excel 시트에 텍스트 박스를 추가합니다.

    텍스트 내용과 함께 폰트, 색상, 정렬 등 다양한 스타일 옵션을 지원하며,
    대시보드의 제목, 설명, 라벨 등을 구성하는 데 적합합니다.

    === 워크북 접근 방법 ===
    - --file-path: 파일 경로로 워크북 열기
    - --use-active: 현재 활성 워크북 사용
    - --workbook-name: 열린 워크북 이름으로 접근 (예: "Sales.xlsx")

    === 텍스트 스타일링 ===
    • --font-size: 글꼴 크기 (포인트)
    • --font-color: 글꼴 색상 (HEX)
    • --font-name: 글꼴 이름
    • --bold, --italic: 굵게, 기울임
    • --alignment: 가로 정렬 (left/center/right)
    • --vertical-alignment: 세로 정렬 (top/middle/bottom)

    === 박스 스타일링 ===
    • --fill-color: 배경 색상
    • --transparency: 배경 투명도
    • --border-color: 테두리 색상
    • --border-width: 테두리 두께
    • --no-border: 테두리 제거

    === 레이아웃 옵션 ===
    • --word-wrap: 자동 줄바꿈 (기본 활성화)
    • --auto-size: 내용에 맞게 크기 자동 조정
    • --left, --top: 위치 지정
    • --width, --height: 크기 지정

    === 대시보드 구성 시나리오 ===

    # 1. 대시보드 메인 제목
    oa excel textbox-add --use-active --text "월별 매출 현황 대시보드" \\
        --left 70 --top 80 --width 760 --height 60 \\
        --font-size 24 --font-color "#FFFFFF" --bold --alignment center \\
        --name "MainTitle"

    # 2. 차트 제목들
    oa excel textbox-add --use-active --text "지역별 매출" \\
        --left 90 --top 180 --width 350 --height 30 \\
        --font-size 16 --font-color "#1D2433" --bold --alignment center \\
        --fill-color "#F8F9FA" --name "Chart1Title"

    oa excel textbox-add --use-active --text "월별 성장률" \\
        --left 460 --top 180 --width 350 --height 30 \\
        --font-size 16 --font-color "#1D2433" --bold --alignment center \\
        --fill-color "#F8F9FA" --name "Chart2Title"

    # 3. 설명 텍스트
    oa excel textbox-add --use-active \\
        --text "이 대시보드는 실시간 매출 데이터를 기반으로 작성되었습니다." \\
        --left 90 --top 520 --width 720 --height 40 \\
        --font-size 10 --font-color "#6C757D" --italic --alignment center \\
        --no-border --name "Description"

    # 4. KPI 라벨들
    oa excel textbox-add --use-active --text "총 매출" \\
        --left 100 --top 400 --width 100 --height 25 \\
        --font-size 14 --font-color "#495057" --bold --alignment center \\
        --fill-color "#E9ECEF" --border-color "#DEE2E6" --border-width 1 \\
        --name "KPI1Label"

    # 5. 다중 라인 텍스트 (자동 줄바꿈)
    oa excel textbox-add --use-active \\
        --text "주요 성과 지표:\\n• 매출 증가율: 15%\\n• 신규 고객: 234명\\n• 고객 만족도: 4.7/5.0" \\
        --left 90 --top 450 --width 300 --height 80 \\
        --font-size 11 --font-color "#495057" --word-wrap --auto-size \\
        --fill-color "#F8F9FA" --border-color "#DEE2E6" --border-width 1 \\
        --name "KPIDetails"

    === 고급 활용 ===

    # 투명 오버레이 텍스트
    oa excel textbox-add --use-active --text "DRAFT" \\
        --left 300 --top 250 --width 200 --height 100 \\
        --font-size 48 --font-color "#DC3545" --bold --alignment center \\
        --vertical-alignment middle --transparency 70 --name "Watermark"

    # 컬러 태그 라벨
    oa excel textbox-add --use-active --text "신규" \\
        --left 450 --top 200 --width 50 --height 20 \\
        --font-size 10 --font-color "#FFFFFF" --bold --alignment center \\
        --vertical-alignment middle --fill-color "#28A745" --no-border \\
        --name "NewTag"

    === 다국어 지원 ===
    • 한글, 영어, 일본어 등 유니코드 텍스트 지원
    • 적절한 폰트 선택으로 가독성 향상
    • 줄바꿈 문자(\\n) 지원

    === 주의사항 ===
    • Windows에서 모든 기능 지원
    • macOS에서는 일부 고급 기능 제한
    • 텍스트가 길 경우 --auto-size 사용 권장
    • 투명도는 0(불투명) - 100(완전투명)
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
                file_path=file_path,
                workbook_name=workbook_name,
                use_active=use_active,
                visible=visible
            )

            # 시트 가져오기
            target_sheet = get_sheet(book, sheet)

            # 텍스트 박스 이름 결정
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
                    raise ValueError(f"텍스트 박스 이름 '{name}'이 이미 존재합니다")
                textbox_name = name
            else:
                textbox_name = generate_unique_shape_name(target_sheet, "TextBox")

            # 텍스트 박스 생성
            try:
                if platform.system() == "Windows":
                    # Windows에서는 COM API 사용
                    textbox = target_sheet.shapes.api.AddTextbox(
                        Orientation=1,  # msoTextOrientationHorizontal
                        Left=left,
                        Top=top,
                        Width=width,
                        Height=height
                    )
                    # xlwings 객체로 래핑
                    textbox_obj = target_sheet.shapes[len(target_sheet.shapes) - 1]
                else:
                    # macOS에서는 xlwings 메서드 사용 (제한적)
                    textbox_obj = target_sheet.shapes.add_textbox(
                        left=left,
                        top=top,
                        width=width,
                        height=height
                    )

                # 텍스트 박스 이름 설정
                textbox_obj.name = textbox_name

                # 텍스트 내용 설정
                # 줄바꿈 문자 처리
                processed_text = text.replace('\\n', '\n')

                if platform.system() == "Windows":
                    textbox_obj.api.TextFrame.Characters().Text = processed_text
                else:
                    # macOS에서는 제한적 지원
                    try:
                        textbox_obj.text = processed_text
                    except:
                        pass

            except Exception as e:
                raise RuntimeError(f"텍스트 박스 생성 실패: {str(e)}")

            # 스타일 적용 (Windows에서만 완전 지원)
            applied_styles = []

            if platform.system() == "Windows":
                try:
                    text_frame = textbox_obj.api.TextFrame
                    characters = text_frame.Characters()

                    # 폰트 설정
                    if font_name:
                        characters.Font.Name = font_name
                        applied_styles.append(f"폰트: {font_name}")

                    if font_size:
                        characters.Font.Size = font_size
                        applied_styles.append(f"크기: {font_size}pt")

                    if font_color:
                        characters.Font.Color = hex_to_rgb(font_color)
                        applied_styles.append(f"글꼴 색상: {font_color}")

                    if bold:
                        characters.Font.Bold = True
                        applied_styles.append("굵게")

                    if italic:
                        characters.Font.Italic = True
                        applied_styles.append("기울임")

                    # 정렬 설정
                    alignment_map = {
                        'left': 1,    # xlLeft
                        'center': 2,  # xlCenter
                        'right': 3    # xlRight
                    }
                    if alignment in alignment_map:
                        text_frame.HorizontalAlignment = alignment_map[alignment]
                        applied_styles.append(f"가로 정렬: {alignment}")

                    vertical_alignment_map = {
                        'top': 1,     # xlTop
                        'middle': 2,  # xlCenter
                        'bottom': 3   # xlBottom
                    }
                    if vertical_alignment in vertical_alignment_map:
                        text_frame.VerticalAlignment = vertical_alignment_map[vertical_alignment]
                        applied_styles.append(f"세로 정렬: {vertical_alignment}")

                    # 텍스트 박스 설정
                    if word_wrap:
                        text_frame.WordWrap = True
                        applied_styles.append("자동 줄바꿈")

                    if auto_size:
                        text_frame.AutoSize = True
                        applied_styles.append("자동 크기 조정")

                    # 배경 색상 및 투명도
                    if fill_color:
                        textbox_obj.api.Fill.ForeColor.RGB = hex_to_rgb(fill_color)
                        applied_styles.append(f"배경 색상: {fill_color}")

                        if transparency is not None:
                            if 0 <= transparency <= 100:
                                textbox_obj.api.Fill.Transparency = transparency / 100.0
                                applied_styles.append(f"배경 투명도: {transparency}%")
                    else:
                        # 배경색이 지정되지 않은 경우 투명하게
                        if transparency is None:
                            textbox_obj.api.Fill.Visible = False
                            applied_styles.append("투명 배경")

                    # 테두리 설정
                    if no_border:
                        textbox_obj.api.Line.Visible = False
                        applied_styles.append("테두리 제거")
                    else:
                        if border_color:
                            textbox_obj.api.Line.Visible = True
                            textbox_obj.api.Line.ForeColor.RGB = hex_to_rgb(border_color)
                            applied_styles.append(f"테두리 색상: {border_color}")

                        if border_width is not None:
                            textbox_obj.api.Line.Visible = True
                            textbox_obj.api.Line.Weight = border_width
                            applied_styles.append(f"테두리 두께: {border_width}pt")

                except Exception as e:
                    applied_styles.append(f"스타일 적용 중 오류: {str(e)}")

            else:
                # macOS에서는 기본 텍스트만 설정
                applied_styles.append("macOS에서는 제한된 스타일링만 지원됩니다")

            # 파일 저장
            if save and file_path:
                book.save()

            # 최종 크기 정보 (자동 크기 조정된 경우)
            final_size = {
                "width": getattr(textbox_obj, 'width', width),
                "height": getattr(textbox_obj, 'height', height)
            }

            # 성공 응답 생성
            response_data = {
                "textbox_name": textbox_name,
                "text_content": processed_text,
                "text_length": len(processed_text),
                "position": {
                    "left": left,
                    "top": top
                },
                "size": final_size,
                "applied_styles": applied_styles,
                "style_count": len([s for s in applied_styles if "오류" not in s]),
                "sheet": target_sheet.name,
                "workbook": normalize_path(book.name),
                "platform_support": "full" if platform.system() == "Windows" else "limited"
            }

            # 텍스트 스타일 요약
            text_style = {
                "font_name": font_name,
                "font_size": font_size,
                "font_color": font_color,
                "alignment": alignment,
                "vertical_alignment": vertical_alignment
            }

            if bold:
                text_style["bold"] = True
            if italic:
                text_style["italic"] = True
            if word_wrap:
                text_style["word_wrap"] = True
            if auto_size:
                text_style["auto_size"] = True

            response_data["text_style"] = text_style

            # 박스 스타일 요약
            box_style = {}
            if fill_color:
                box_style["fill_color"] = fill_color
            if transparency is not None:
                box_style["transparency"] = transparency
            if border_color:
                box_style["border_color"] = border_color
            if border_width is not None:
                box_style["border_width"] = border_width
            if no_border:
                box_style["no_border"] = True

            if box_style:
                response_data["box_style"] = box_style

            message = f"텍스트 박스 '{textbox_name}'이 성공적으로 생성되었습니다"

            response = create_success_response(
                data=response_data,
                command="textbox-add",
                message=message,
                execution_time_ms=timer.execution_time_ms,
                book=book,
                text_length=len(processed_text),
                styles_applied=len(applied_styles)
            )

            print(json.dumps(response, ensure_ascii=False, indent=2))

    except Exception as e:
        error_response = create_error_response(e, "textbox-add")
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


if __name__ == '__main__':
    textbox_add()