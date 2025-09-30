"""
PowerPoint 도형 추가 명령어
슬라이드에 도형을 추가합니다.
"""

import json
from pathlib import Path
from typing import Optional

import typer
from pptx import Presentation
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.util import Inches

from pyhub_office_automation.version import get_version

from .utils import (
    ShapeType,
    create_error_response,
    create_success_response,
    normalize_path,
    parse_color,
    validate_slide_number,
)


def content_add_shape(
    file_path: str = typer.Option(..., "--file-path", help="PowerPoint 파일 경로"),
    slide_number: int = typer.Option(..., "--slide-number", help="도형을 추가할 슬라이드 번호 (1부터 시작)"),
    shape_type: str = typer.Option(..., "--shape-type", help="도형 유형 (rectangle/ellipse/star/arrow 등)"),
    left: float = typer.Option(..., "--left", help="도형 왼쪽 위치 (인치)"),
    top: float = typer.Option(..., "--top", help="도형 상단 위치 (인치)"),
    width: float = typer.Option(..., "--width", help="도형 너비 (인치)"),
    height: float = typer.Option(..., "--height", help="도형 높이 (인치)"),
    fill_color: Optional[str] = typer.Option(None, "--fill-color", help="채우기 색상 (색상명 또는 #RGB/#RRGGBB)"),
    line_color: Optional[str] = typer.Option(None, "--line-color", help="테두리 색상 (색상명 또는 #RGB/#RRGGBB)"),
    line_width: Optional[float] = typer.Option(None, "--line-width", help="테두리 두께 (포인트)"),
    text: Optional[str] = typer.Option(None, "--text", help="도형 내부 텍스트"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 (json/text)"),
):
    """
    PowerPoint 슬라이드에 도형을 추가합니다.

    지원 도형 유형:
      - rectangle: 사각형
      - rounded-rectangle: 둥근 사각형
      - ellipse: 타원
      - arrow-right: 오른쪽 화살표
      - arrow-left: 왼쪽 화살표
      - arrow-up: 위쪽 화살표
      - arrow-down: 아래쪽 화살표
      - star: 별
      - pentagon: 오각형
      - hexagon: 육각형

    예제:
        oa ppt content-add-shape --file-path "presentation.pptx" --slide-number 1 --shape-type rectangle --left 1 --top 2 --width 3 --height 2 --fill-color blue
        oa ppt content-add-shape --file-path "presentation.pptx" --slide-number 2 --shape-type star --left 2 --top 3 --width 1.5 --height 1.5 --fill-color "#FFD700" --text "중요"
        oa ppt content-add-shape --file-path "presentation.pptx" --slide-number 3 --shape-type arrow-right --left 1 --top 1 --width 2 --height 1 --fill-color red --line-color black --line-width 2
    """
    try:
        # 도형 유형 검증
        shape_type_map = {
            ShapeType.RECTANGLE: MSO_AUTO_SHAPE_TYPE.RECTANGLE,
            ShapeType.ROUNDED_RECTANGLE: MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
            ShapeType.ELLIPSE: MSO_AUTO_SHAPE_TYPE.OVAL,
            ShapeType.ARROW_RIGHT: MSO_AUTO_SHAPE_TYPE.RIGHT_ARROW,
            ShapeType.ARROW_LEFT: MSO_AUTO_SHAPE_TYPE.LEFT_ARROW,
            ShapeType.ARROW_UP: MSO_AUTO_SHAPE_TYPE.UP_ARROW,
            ShapeType.ARROW_DOWN: MSO_AUTO_SHAPE_TYPE.DOWN_ARROW,
            ShapeType.STAR: MSO_AUTO_SHAPE_TYPE.STAR_5,
            ShapeType.PENTAGON: MSO_AUTO_SHAPE_TYPE.PENTAGON,
            ShapeType.HEXAGON: MSO_AUTO_SHAPE_TYPE.HEXAGON,
        }

        if shape_type not in shape_type_map:
            available_types = ", ".join(shape_type_map.keys())
            raise ValueError(f"지원하지 않는 도형 유형: {shape_type}. 사용 가능: {available_types}")

        # 파일 경로 정규화 및 존재 확인
        normalized_path = normalize_path(file_path)
        pptx_path = Path(normalized_path).resolve()

        if not pptx_path.exists():
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {pptx_path}")

        # 프레젠테이션 열기
        prs = Presentation(str(pptx_path))

        # 슬라이드 번호 검증
        slide_idx = validate_slide_number(slide_number, len(prs.slides))
        slide = prs.slides[slide_idx]

        # 도형 추가
        mso_shape_type = shape_type_map[shape_type]
        shape = slide.shapes.add_shape(mso_shape_type, Inches(left), Inches(top), Inches(width), Inches(height))

        # 채우기 색상 설정
        if fill_color is not None:
            color = parse_color(fill_color)
            shape.fill.solid()
            shape.fill.fore_color.rgb = color

        # 테두리 설정
        if line_color is not None:
            color = parse_color(line_color)
            shape.line.color.rgb = color

        if line_width is not None:
            from pptx.util import Pt

            shape.line.width = Pt(line_width)

        # 텍스트 추가
        if text is not None:
            if hasattr(shape, "text_frame"):
                text_frame = shape.text_frame
                text_frame.clear()
                paragraph = text_frame.paragraphs[0]
                run = paragraph.add_run()
                run.text = text

        # 저장
        prs.save(str(pptx_path))

        # 결과 데이터 구성
        result_data = {
            "file": str(pptx_path),
            "slide_number": slide_number,
            "shape_type": shape_type,
            "position": {
                "left": left,
                "top": top,
                "width": width,
                "height": height,
            },
        }

        if fill_color is not None:
            result_data["fill_color"] = fill_color

        if line_color is not None:
            result_data["line_color"] = line_color

        if line_width is not None:
            result_data["line_width"] = line_width

        if text is not None:
            result_data["text"] = text

        # 성공 응답
        message = f"슬라이드 {slide_number}에 도형 '{shape_type}'을(를) 추가했습니다"

        response = create_success_response(
            data=result_data,
            command="content-add-shape",
            message=message,
        )

        # 출력
        if output_format == "json":
            typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
        else:
            typer.echo(f"✅ {message}")
            typer.echo(f"📄 파일: {pptx_path.name}")
            typer.echo(f"📍 슬라이드: {slide_number}")
            typer.echo(f"🔷 도형: {shape_type}")
            typer.echo(f"📐 위치: {left}in × {top}in")
            typer.echo(f"📏 크기: {width}in × {height}in")
            if fill_color:
                typer.echo(f"🎨 채우기: {fill_color}")
            if line_color:
                typer.echo(f"✏️ 테두리: {line_color}")
            if text:
                typer.echo(f"📝 텍스트: {text}")

    except FileNotFoundError as e:
        error_response = create_error_response(e, "content-add-shape")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)

    except ValueError as e:
        error_response = create_error_response(e, "content-add-shape")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)

    except Exception as e:
        error_response = create_error_response(e, "content-add-shape")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ 예기치 않은 오류: {str(e)}", err=True)
        raise typer.Exit(1)


if __name__ == "__main__":
    typer.run(content_add_shape)
