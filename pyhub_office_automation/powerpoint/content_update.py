"""
PowerPoint 콘텐츠 업데이트 명령어
슬라이드의 기존 콘텐츠(텍스트, 이미지, 위치 등)를 수정합니다.
"""

import json
from pathlib import Path
from typing import Optional

import typer
from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches

from pyhub_office_automation.version import get_version

from .utils import (
    create_error_response,
    create_success_response,
    get_shape_by_index_or_name,
    normalize_path,
    validate_slide_number,
)


def content_update(
    file_path: str = typer.Option(..., "--file-path", help="PowerPoint 파일 경로"),
    slide_number: int = typer.Option(..., "--slide-number", help="콘텐츠를 업데이트할 슬라이드 번호 (1부터 시작)"),
    shape_index: Optional[int] = typer.Option(None, "--shape-index", help="Shape 인덱스 (0부터 시작)"),
    shape_name: Optional[str] = typer.Option(None, "--shape-name", help="Shape 이름"),
    text: Optional[str] = typer.Option(None, "--text", help="업데이트할 텍스트 내용"),
    image_path: Optional[str] = typer.Option(None, "--image-path", help="교체할 이미지 파일 경로"),
    left: Optional[float] = typer.Option(None, "--left", help="새 위치 - 왼쪽 (인치)"),
    top: Optional[float] = typer.Option(None, "--top", help="새 위치 - 상단 (인치)"),
    width: Optional[float] = typer.Option(None, "--width", help="새 크기 - 너비 (인치)"),
    height: Optional[float] = typer.Option(None, "--height", help="새 크기 - 높이 (인치)"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 (json/text)"),
):
    """
    PowerPoint 슬라이드의 기존 콘텐츠를 업데이트합니다.

    Shape 선택 (둘 중 하나만 지정):
      --shape-index: Shape 인덱스 (0부터 시작)
      --shape-name: Shape 이름

    업데이트 옵션:
      --text: 텍스트 내용 변경
      --image-path: 이미지 교체 (picture shape만 가능)
      --left, --top: 위치 이동
      --width, --height: 크기 조정

    예제:
        oa ppt content-update --file-path "presentation.pptx" --slide-number 1 --shape-index 0 --text "새 제목"
        oa ppt content-update --file-path "presentation.pptx" --slide-number 2 --shape-name "Picture 1" --image-path "new_image.png"
        oa ppt content-update --file-path "presentation.pptx" --slide-number 3 --shape-index 1 --left 2 --top 2 --width 4 --height 3
    """
    try:
        # 입력 검증
        if shape_index is None and shape_name is None:
            raise ValueError("--shape-index 또는 --shape-name 중 하나는 반드시 지정해야 합니다")

        if shape_index is not None and shape_name is not None:
            raise ValueError("--shape-index와 --shape-name은 동시에 사용할 수 없습니다")

        if text is None and image_path is None and left is None and top is None and width is None and height is None:
            raise ValueError(
                "업데이트할 내용을 지정해야 합니다 (--text, --image-path, --left, --top, --width, --height 중 하나 이상)"
            )

        # 파일 경로 정규화 및 존재 확인
        normalized_pptx_path = normalize_path(file_path)
        pptx_path = Path(normalized_pptx_path).resolve()

        if not pptx_path.exists():
            raise FileNotFoundError(f"PowerPoint 파일을 찾을 수 없습니다: {pptx_path}")

        # 이미지 경로 검증
        if image_path:
            normalized_image_path = normalize_path(image_path)
            img_path = Path(normalized_image_path).resolve()
            if not img_path.exists():
                raise FileNotFoundError(f"이미지 파일을 찾을 수 없습니다: {img_path}")
        else:
            img_path = None

        # 프레젠테이션 열기
        prs = Presentation(str(pptx_path))

        # 슬라이드 번호 검증
        slide_idx = validate_slide_number(slide_number, len(prs.slides))
        slide = prs.slides[slide_idx]

        # Shape 찾기
        identifier = shape_index if shape_index is not None else shape_name
        shape = get_shape_by_index_or_name(slide, identifier)

        # 업데이트 내용 기록
        updates = []

        # 텍스트 업데이트
        if text is not None:
            if not hasattr(shape, "text_frame"):
                raise ValueError(f"이 Shape는 텍스트를 지원하지 않습니다: {shape.shape_type}")

            shape.text_frame.clear()
            paragraph = shape.text_frame.paragraphs[0]
            run = paragraph.add_run()
            run.text = text
            updates.append(f"텍스트 변경: '{text[:50]}...' ({len(text)}자)")

        # 이미지 교체
        if image_path:
            if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
                raise ValueError(f"이미지 교체는 picture shape만 가능합니다 (현재: {shape.shape_type})")

            # 기존 이미지 정보 저장
            old_left = shape.left
            old_top = shape.top
            old_width = shape.width
            old_height = shape.height

            # Shape 인덱스 찾기 (삭제 후 같은 위치에 추가하기 위해)
            shape_idx = None
            for idx, s in enumerate(slide.shapes):
                if s == shape:
                    shape_idx = idx
                    break

            # 기존 shape 삭제
            sp = shape.element
            sp.getparent().remove(sp)

            # 새 이미지 추가 (같은 위치와 크기)
            new_left = Inches(left) if left is not None else old_left
            new_top = Inches(top) if top is not None else old_top
            new_width = Inches(width) if width is not None else old_width
            new_height = Inches(height) if height is not None else old_height

            picture = slide.shapes.add_picture(str(img_path), new_left, new_top, width=new_width, height=new_height)

            # 새로 추가된 shape를 참조
            shape = picture
            updates.append(f"이미지 교체: {img_path.name}")

        # 위치 업데이트
        if left is not None or top is not None:
            if left is not None:
                shape.left = Inches(left)
                updates.append(f"위치 변경 (left): {left}in")
            if top is not None:
                shape.top = Inches(top)
                updates.append(f"위치 변경 (top): {top}in")

        # 크기 업데이트
        if width is not None or height is not None:
            if width is not None:
                shape.width = Inches(width)
                updates.append(f"크기 변경 (width): {width}in")
            if height is not None:
                shape.height = Inches(height)
                updates.append(f"크기 변경 (height): {height}in")

        # 저장
        prs.save(str(pptx_path))

        # 결과 데이터 구성
        result_data = {
            "file": str(pptx_path),
            "slide_number": slide_number,
            "shape_identifier": shape_name if shape_name else f"index_{shape_index}",
            "shape_type": str(shape.shape_type),
            "updates": updates,
            "update_count": len(updates),
        }

        # 현재 위치/크기 정보
        result_data["current_position"] = {
            "left": round(shape.left / 914400, 2),  # EMU to inches
            "top": round(shape.top / 914400, 2),
            "width": round(shape.width / 914400, 2),
            "height": round(shape.height / 914400, 2),
        }

        # 성공 응답
        message = f"슬라이드 {slide_number}의 shape를 업데이트했습니다 ({len(updates)}개 항목)"

        response = create_success_response(
            data=result_data,
            command="content-update",
            message=message,
        )

        # 출력
        if output_format == "json":
            typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
        else:
            typer.echo(f"✅ {message}")
            typer.echo(f"📄 파일: {pptx_path.name}")
            typer.echo(f"📍 슬라이드: {slide_number}")
            typer.echo(f"🎯 Shape: {result_data['shape_identifier']}")
            typer.echo(f"📦 Shape 타입: {result_data['shape_type']}")
            typer.echo(f"\n🔄 업데이트 내역:")
            for update in updates:
                typer.echo(f"  • {update}")
            typer.echo(f"\n📐 현재 위치/크기:")
            pos = result_data["current_position"]
            typer.echo(f"  위치: {pos['left']}in × {pos['top']}in")
            typer.echo(f"  크기: {pos['width']}in × {pos['height']}in")

    except FileNotFoundError as e:
        error_response = create_error_response(e, "content-update")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)

    except ValueError as e:
        error_response = create_error_response(e, "content-update")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)

    except Exception as e:
        error_response = create_error_response(e, "content-update")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ 예기치 않은 오류: {str(e)}", err=True)
        raise typer.Exit(1)


if __name__ == "__main__":
    typer.run(content_update)
