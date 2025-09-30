"""
PowerPoint 노트 내보내기 명령어 (COM-First)
슬라이드 노트(발표자 노트)를 텍스트 또는 JSON 형식으로 추출합니다.
"""

import json
from pathlib import Path
from typing import Optional

import typer

from pyhub_office_automation.version import get_version

from .utils import (
    PowerPointBackend,
    create_error_response,
    create_success_response,
    get_or_open_presentation,
    get_powerpoint_backend,
    normalize_path,
)


def export_notes(
    output_path: str = typer.Option(..., "--output-path", help="노트 저장 경로 (.txt 또는 .json)"),
    slides: Optional[str] = typer.Option(None, "--slides", help="내보낼 슬라이드 범위 (예: '1-5', '1,3,5', 'all')"),
    include_slide_titles: bool = typer.Option(True, "--include-titles/--no-titles", help="슬라이드 제목 포함 (기본: True)"),
    separator: str = typer.Option("\n\n" + "=" * 50 + "\n\n", "--separator", help="슬라이드 구분자"),
    file_path: Optional[str] = typer.Option(None, "--file-path", help="PowerPoint 파일 경로"),
    presentation_name: Optional[str] = typer.Option(None, "--presentation-name", help="열려있는 프레젠테이션 이름 (COM 전용)"),
    backend: str = typer.Option("auto", "--backend", help="백엔드 선택 (auto/com/python-pptx)"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 (json/text)"),
):
    """
    PowerPoint 슬라이드 노트를 추출합니다.

    COM-First: Windows에서는 COM 백엔드 우선, python-pptx는 fallback

    **백엔드 선택**:
    - auto (기본): 자동으로 최적 백엔드 선택 (Windows COM 우선)
    - com: Windows COM 강제 사용 (완전한 기능)
    - python-pptx: python-pptx 강제 사용 (기본 기능)

    **COM 백엔드 (Windows)**:
    - ✅ Slide.NotesPage.Shapes(2).TextFrame.TextRange.Text 사용
    - ✅ 슬라이드 제목 추출
    - ✅ 슬라이드 범위 선택

    **python-pptx 백엔드**:
    - ✅ slide.notes_slide.notes_text_frame.text 사용
    - ✅ 슬라이드 제목 추출
    - ✅ 파일 저장 필수 (--file-path 필수)

    **출력 형식**:
    - .txt: 일반 텍스트 파일 (구분자로 슬라이드 구분)
    - .json: JSON 형식 (슬라이드 번호, 제목, 노트 구조화)

    예제:
        # COM 백엔드 (활성 프레젠테이션 전체, TXT)
        oa ppt export-notes --output-path "notes.txt"

        # 특정 슬라이드만 (1-10번, JSON)
        oa ppt export-notes --output-path "notes.json" --slides "1-10"

        # 제목 제외
        oa ppt export-notes --output-path "notes_only.txt" --no-titles

        # 커스텀 구분자
        oa ppt export-notes --output-path "notes.txt" --separator "\\n---\\n" --presentation-name "report.pptx"
    """

    try:
        # 출력 경로 검증
        normalized_output_path = normalize_path(output_path)
        notes_path = Path(normalized_output_path).resolve()

        # 디렉토리 생성
        notes_path.parent.mkdir(parents=True, exist_ok=True)

        # 파일 확장자로 출력 형식 결정
        file_ext = notes_path.suffix.lower()
        if file_ext not in [".txt", ".json"]:
            result = create_error_response(
                command="export-notes",
                error=f"지원하지 않는 파일 형식: {file_ext}. .txt 또는 .json을 사용하세요.",
                error_type="ValueError",
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

        export_format = "json" if file_ext == ".json" else "text"

        # 백엔드 결정
        try:
            selected_backend = get_powerpoint_backend(force_backend=backend if backend != "auto" else None)
        except (ValueError, RuntimeError) as e:
            result = create_error_response(
                command="export-notes",
                error=str(e),
                error_type=type(e).__name__,
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

        # 프레젠테이션 가져오기
        try:
            backend_inst, prs = get_or_open_presentation(
                file_path=file_path,
                presentation_name=presentation_name,
                backend=selected_backend,
            )
        except Exception as e:
            result = create_error_response(
                command="export-notes",
                error=f"프레젠테이션을 열 수 없습니다: {str(e)}",
                error_type=type(e).__name__,
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

        # 백엔드별 처리
        if selected_backend == PowerPointBackend.COM.value:
            # COM 백엔드
            try:
                total_slides = prs.Slides.Count

                # 슬라이드 범위 파싱
                if slides and slides.lower() != "all":
                    from .export_pdf import parse_slide_range

                    slide_numbers = parse_slide_range(slides, total_slides)

                    if not slide_numbers:
                        result = create_error_response(
                            command="export-notes",
                            error=f"유효하지 않은 슬라이드 범위: {slides}",
                            error_type="ValueError",
                        )
                        typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
                        raise typer.Exit(1)
                else:
                    slide_numbers = list(range(1, total_slides + 1))

                # 노트 추출
                notes_data = []
                notes_with_content = 0

                for slide_num in slide_numbers:
                    slide = prs.Slides(slide_num)

                    # 슬라이드 제목 추출
                    slide_title = ""
                    if include_slide_titles:
                        try:
                            # 첫 번째 shape가 제목인 경우가 많음
                            for shape in slide.Shapes:
                                if hasattr(shape, "TextFrame") and hasattr(shape.TextFrame, "TextRange"):
                                    text = shape.TextFrame.TextRange.Text.strip()
                                    if text:
                                        slide_title = text
                                        break
                        except Exception:
                            pass

                    # 노트 추출
                    # NotesPage.Shapes(2) = 노트 텍스트 프레임 (1은 슬라이드 미리보기)
                    notes_text = ""
                    try:
                        notes_page = slide.NotesPage
                        if notes_page.Shapes.Count >= 2:
                            notes_shape = notes_page.Shapes(2)  # COM은 1-based
                            if hasattr(notes_shape, "TextFrame"):
                                notes_text = notes_shape.TextFrame.TextRange.Text.strip()
                    except Exception:
                        pass

                    if notes_text:
                        notes_with_content += 1

                    notes_data.append(
                        {
                            "slide_number": slide_num,
                            "slide_title": slide_title,
                            "notes": notes_text,
                        }
                    )

                # 파일로 저장
                if export_format == "json":
                    # JSON 형식
                    output_data = {
                        "presentation": {
                            "total_slides": total_slides,
                            "exported_count": len(notes_data),
                            "notes_count": notes_with_content,
                        },
                        "slides": notes_data,
                    }

                    with open(notes_path, "w", encoding="utf-8") as f:
                        json.dump(output_data, f, ensure_ascii=False, indent=2)

                else:
                    # 텍스트 형식
                    text_lines = []

                    for note_item in notes_data:
                        slide_num = note_item["slide_number"]
                        slide_title = note_item["slide_title"]
                        notes_text = note_item["notes"]

                        # 슬라이드 헤더
                        text_lines.append(f"슬라이드 {slide_num}")

                        if include_slide_titles and slide_title:
                            text_lines.append(f"제목: {slide_title}")

                        text_lines.append("")

                        # 노트 내용
                        if notes_text:
                            text_lines.append(notes_text)
                        else:
                            text_lines.append("(노트 없음)")

                        text_lines.append(separator)

                    with open(notes_path, "w", encoding="utf-8") as f:
                        f.write("\n".join(text_lines))

                # 파일 크기
                file_size_kb = notes_path.stat().st_size / 1024

                # 성공 응답
                result_data = {
                    "backend": "com",
                    "output_file": str(notes_path),
                    "output_file_name": notes_path.name,
                    "file_size_kb": round(file_size_kb, 2),
                    "export_format": export_format,
                    "total_slides": total_slides,
                    "exported_count": len(notes_data),
                    "notes_with_content": notes_with_content,
                    "include_titles": include_slide_titles,
                }

                message = f"노트 내보내기 완료 (COM): {len(notes_data)}개 슬라이드"

            except Exception as e:
                result = create_error_response(
                    command="export-notes",
                    error=f"노트 내보내기 실패: {str(e)}",
                    error_type=type(e).__name__,
                )
                typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
                raise typer.Exit(1)

        else:
            # python-pptx 백엔드
            if not file_path:
                result = create_error_response(
                    command="export-notes",
                    error="python-pptx 백엔드는 --file-path 옵션이 필수입니다",
                    error_type="ValueError",
                )
                typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
                raise typer.Exit(1)

            try:
                total_slides = len(prs.slides)

                # 슬라이드 범위 파싱
                if slides and slides.lower() != "all":
                    from .export_pdf import parse_slide_range

                    slide_numbers = parse_slide_range(slides, total_slides)

                    if not slide_numbers:
                        result = create_error_response(
                            command="export-notes",
                            error=f"유효하지 않은 슬라이드 범위: {slides}",
                            error_type="ValueError",
                        )
                        typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
                        raise typer.Exit(1)

                    # python-pptx는 0-based 인덱싱
                    slide_indices = [num - 1 for num in slide_numbers]
                else:
                    slide_indices = list(range(total_slides))

                # 노트 추출
                notes_data = []
                notes_with_content = 0

                for idx in slide_indices:
                    slide = prs.slides[idx]
                    slide_num = idx + 1

                    # 슬라이드 제목 추출
                    slide_title = ""
                    if include_slide_titles:
                        try:
                            # 첫 번째 shape가 제목인 경우가 많음
                            for shape in slide.shapes:
                                if hasattr(shape, "text") and shape.text.strip():
                                    slide_title = shape.text.strip()
                                    break
                        except Exception:
                            pass

                    # 노트 추출
                    notes_text = ""
                    try:
                        if slide.has_notes_slide:
                            notes_slide = slide.notes_slide
                            if notes_slide.notes_text_frame:
                                notes_text = notes_slide.notes_text_frame.text.strip()
                    except Exception:
                        pass

                    if notes_text:
                        notes_with_content += 1

                    notes_data.append(
                        {
                            "slide_number": slide_num,
                            "slide_title": slide_title,
                            "notes": notes_text,
                        }
                    )

                # 파일로 저장
                if export_format == "json":
                    # JSON 형식
                    output_data = {
                        "presentation": {
                            "total_slides": total_slides,
                            "exported_count": len(notes_data),
                            "notes_count": notes_with_content,
                        },
                        "slides": notes_data,
                    }

                    with open(notes_path, "w", encoding="utf-8") as f:
                        json.dump(output_data, f, ensure_ascii=False, indent=2)

                else:
                    # 텍스트 형식
                    text_lines = []

                    for note_item in notes_data:
                        slide_num = note_item["slide_number"]
                        slide_title = note_item["slide_title"]
                        notes_text = note_item["notes"]

                        # 슬라이드 헤더
                        text_lines.append(f"슬라이드 {slide_num}")

                        if include_slide_titles and slide_title:
                            text_lines.append(f"제목: {slide_title}")

                        text_lines.append("")

                        # 노트 내용
                        if notes_text:
                            text_lines.append(notes_text)
                        else:
                            text_lines.append("(노트 없음)")

                        text_lines.append(separator)

                    with open(notes_path, "w", encoding="utf-8") as f:
                        f.write("\n".join(text_lines))

                # 파일 크기
                file_size_kb = notes_path.stat().st_size / 1024

                # 성공 응답
                result_data = {
                    "backend": "python-pptx",
                    "output_file": str(notes_path),
                    "output_file_name": notes_path.name,
                    "file_size_kb": round(file_size_kb, 2),
                    "export_format": export_format,
                    "total_slides": total_slides,
                    "exported_count": len(notes_data),
                    "notes_with_content": notes_with_content,
                    "include_titles": include_slide_titles,
                }

                message = f"노트 내보내기 완료 (python-pptx): {len(notes_data)}개 슬라이드"

            except Exception as e:
                result = create_error_response(
                    command="export-notes",
                    error=f"노트 내보내기 실패: {str(e)}",
                    error_type=type(e).__name__,
                )
                typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
                raise typer.Exit(1)

        # 성공 응답
        response = create_success_response(
            data=result_data,
            command="export-notes",
            message=message,
        )

        # 출력
        if output_format == "json":
            typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
        else:
            typer.echo(f"✅ {message}")
            typer.echo(f"📄 파일: {notes_path}")
            typer.echo(f"💾 크기: {result_data['file_size_kb']} KB")
            typer.echo(f"📊 슬라이드: {result_data['exported_count']}개 / 총 {result_data['total_slides']}개")
            typer.echo(f"📝 노트 있음: {result_data['notes_with_content']}개")
            typer.echo(f"📋 형식: {export_format.upper()}")
            if include_slide_titles:
                typer.echo("📌 슬라이드 제목: 포함")

    except typer.Exit:
        raise
    except Exception as e:
        result = create_error_response(
            command="export-notes",
            error=str(e),
            error_type=type(e).__name__,
        )
        typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
        raise typer.Exit(1)
    finally:
        # python-pptx는 자동 정리, COM은 유지
        pass


if __name__ == "__main__":
    typer.run(export_notes)
