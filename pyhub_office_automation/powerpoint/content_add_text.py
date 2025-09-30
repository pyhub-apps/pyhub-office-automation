"""
PowerPoint 텍스트 추가 명령어
슬라이드에 텍스트를 추가합니다 (플레이스홀더 또는 자유 위치).
"""

import json
from pathlib import Path
from typing import Optional

import typer
from pptx import Presentation
from pptx.util import Inches, Pt

from pyhub_office_automation.version import get_version

from .utils import (
    PlaceholderType,
    create_error_response,
    create_success_response,
    get_placeholder_by_type,
    normalize_path,
    parse_color,
    validate_slide_number,
)


def content_add_text(
    file_path: str = typer.Option(..., "--file-path", help="PowerPoint 파일 경로"),
    slide_number: int = typer.Option(..., "--slide-number", help="텍스트를 추가할 슬라이드 번호 (1부터 시작)"),
    placeholder: Optional[str] = typer.Option(
        None, "--placeholder", help="플레이스홀더 유형 (title/body/subtitle) - 이 옵션 사용 시 위치 옵션 무시"
    ),
    text: Optional[str] = typer.Option(None, "--text", help="추가할 텍스트 (직접 입력)"),
    text_file: Optional[str] = typer.Option(None, "--text-file", help="텍스트 파일 경로 (.txt)"),
    left: Optional[float] = typer.Option(None, "--left", help="텍스트 박스 왼쪽 위치 (인치)"),
    top: Optional[float] = typer.Option(None, "--top", help="텍스트 박스 상단 위치 (인치)"),
    width: Optional[float] = typer.Option(3.0, "--width", help="텍스트 박스 너비 (인치, 기본값: 3.0)"),
    height: Optional[float] = typer.Option(1.0, "--height", help="텍스트 박스 높이 (인치, 기본값: 1.0)"),
    font_size: Optional[int] = typer.Option(None, "--font-size", help="글꼴 크기 (포인트)"),
    font_color: Optional[str] = typer.Option(None, "--font-color", help="글꼴 색상 (색상명 또는 #RGB/#RRGGBB)"),
    bold: bool = typer.Option(False, "--bold", help="굵게 적용"),
    italic: bool = typer.Option(False, "--italic", help="기울임꼴 적용"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 (json/text)"),
):
    """
    PowerPoint 슬라이드에 텍스트를 추가합니다.

    플레이스홀더 모드 (--placeholder 지정):
      title, body, subtitle 중 하나를 지정하면 해당 플레이스홀더에 텍스트 추가

    자유 위치 모드 (--left, --top 지정):
      지정된 위치에 텍스트 박스를 생성하여 텍스트 추가

    텍스트 입력:
      --text: 직접 텍스트 입력
      --text-file: 파일에서 텍스트 읽기 (.txt)

    예제:
        oa ppt content-add-text --file-path "presentation.pptx" --slide-number 1 --placeholder title --text "제목"
        oa ppt content-add-text --file-path "presentation.pptx" --slide-number 2 --left 1 --top 2 --text "본문" --font-size 18
        oa ppt content-add-text --file-path "presentation.pptx" --slide-number 3 --placeholder body --text-file "content.txt"
    """
    try:
        # 입력 검증
        if not text and not text_file:
            raise ValueError("--text 또는 --text-file 중 하나는 반드시 지정해야 합니다")

        if text and text_file:
            raise ValueError("--text와 --text-file은 동시에 사용할 수 없습니다")

        if placeholder and (left is not None or top is not None):
            raise ValueError("--placeholder와 --left/--top은 동시에 사용할 수 없습니다")

        if not placeholder and (left is None or top is None):
            raise ValueError("--placeholder를 지정하지 않은 경우 --left와 --top을 모두 지정해야 합니다")

        if placeholder and placeholder not in [PlaceholderType.TITLE, PlaceholderType.BODY, PlaceholderType.SUBTITLE]:
            raise ValueError(f"잘못된 플레이스홀더 유형: {placeholder}. 사용 가능: title, body, subtitle")

        # 파일 경로 정규화 및 존재 확인
        normalized_path = normalize_path(file_path)
        pptx_path = Path(normalized_path).resolve()

        if not pptx_path.exists():
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {pptx_path}")

        # 텍스트 로드
        text_content = None
        if text_file:
            text_file_path = Path(normalize_path(text_file)).resolve()
            if not text_file_path.exists():
                raise FileNotFoundError(f"텍스트 파일을 찾을 수 없습니다: {text_file_path}")
            with open(text_file_path, "r", encoding="utf-8") as f:
                text_content = f.read()
        else:
            text_content = text

        # 프레젠테이션 열기
        prs = Presentation(str(pptx_path))

        # 슬라이드 번호 검증
        slide_idx = validate_slide_number(slide_number, len(prs.slides))
        slide = prs.slides[slide_idx]

        # 텍스트 추가 처리
        text_frame = None
        shape = None
        mode = "placeholder" if placeholder else "position"

        if placeholder:
            # 플레이스홀더 모드
            shape = get_placeholder_by_type(slide, placeholder)
            if shape is None:
                raise ValueError(f"슬라이드 {slide_number}에 '{placeholder}' 플레이스홀더가 없습니다")
            text_frame = shape.text_frame
        else:
            # 자유 위치 모드
            shape = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
            text_frame = shape.text_frame

        # 텍스트 설정
        text_frame.clear()  # 기존 내용 제거
        paragraph = text_frame.paragraphs[0]
        run = paragraph.add_run()
        run.text = text_content

        # 스타일 적용
        if font_size is not None:
            run.font.size = Pt(font_size)

        if font_color is not None:
            color = parse_color(font_color)
            run.font.color.rgb = color

        if bold:
            run.font.bold = True

        if italic:
            run.font.italic = True

        # 저장
        prs.save(str(pptx_path))

        # 결과 데이터 구성
        result_data = {
            "file": str(pptx_path),
            "slide_number": slide_number,
            "mode": mode,
            "text_length": len(text_content),
            "text_preview": text_content[:100] + "..." if len(text_content) > 100 else text_content,
        }

        if placeholder:
            result_data["placeholder"] = placeholder
        else:
            result_data["position"] = {
                "left": left,
                "top": top,
                "width": width,
                "height": height,
            }

        if font_size is not None:
            result_data["font_size"] = font_size

        if font_color is not None:
            result_data["font_color"] = font_color

        result_data["bold"] = bold
        result_data["italic"] = italic

        # 성공 응답
        message = f"슬라이드 {slide_number}에 텍스트를 추가했습니다"
        if placeholder:
            message += f" (플레이스홀더: {placeholder})"
        else:
            message += f" (위치: {left}in, {top}in)"

        response = create_success_response(
            data=result_data,
            command="content-add-text",
            message=message,
        )

        # 출력
        if output_format == "json":
            typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
        else:
            typer.echo(f"✅ {message}")
            typer.echo(f"📄 파일: {pptx_path.name}")
            typer.echo(f"📍 슬라이드: {slide_number}")
            if placeholder:
                typer.echo(f"🎯 플레이스홀더: {placeholder}")
            else:
                typer.echo(f"📐 위치: {left}in × {top}in")
                typer.echo(f"📏 크기: {width}in × {height}in")
            typer.echo(f"📝 텍스트 길이: {len(text_content)}자")
            typer.echo(f"📄 미리보기: {result_data['text_preview']}")

    except FileNotFoundError as e:
        error_response = create_error_response(e, "content-add-text")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)

    except ValueError as e:
        error_response = create_error_response(e, "content-add-text")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)

    except Exception as e:
        error_response = create_error_response(e, "content-add-text")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ 예기치 않은 오류: {str(e)}", err=True)
        raise typer.Exit(1)


if __name__ == "__main__":
    typer.run(content_add_text)
