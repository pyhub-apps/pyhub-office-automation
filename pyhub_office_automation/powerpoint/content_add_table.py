"""
PowerPoint 표 추가 명령어
슬라이드에 표를 추가하고 데이터를 채웁니다.
"""

import csv
import json
from pathlib import Path
from typing import Optional

import typer
from pptx import Presentation
from pptx.util import Inches

from pyhub_office_automation.version import get_version

from .utils import create_error_response, create_success_response, normalize_path, validate_slide_number


def content_add_table(
    file_path: str = typer.Option(..., "--file-path", help="PowerPoint 파일 경로"),
    slide_number: int = typer.Option(..., "--slide-number", help="표를 추가할 슬라이드 번호 (1부터 시작)"),
    rows: int = typer.Option(..., "--rows", help="표 행 수"),
    cols: int = typer.Option(..., "--cols", help="표 열 수"),
    left: float = typer.Option(..., "--left", help="표 왼쪽 위치 (인치)"),
    top: float = typer.Option(..., "--top", help="표 상단 위치 (인치)"),
    width: float = typer.Option(..., "--width", help="표 너비 (인치)"),
    height: float = typer.Option(..., "--height", help="표 높이 (인치)"),
    data: Optional[str] = typer.Option(None, "--data", help="표 데이터 (JSON 2차원 배열)"),
    data_file: Optional[str] = typer.Option(None, "--data-file", help="표 데이터 파일 (.csv 또는 .json)"),
    first_row_header: bool = typer.Option(False, "--first-row-header", help="첫 행을 헤더로 처리"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 (json/text)"),
):
    """
    PowerPoint 슬라이드에 표를 추가하고 데이터를 채웁니다.

    데이터 입력 방법:
      --data: JSON 2차원 배열로 직접 입력 (예: '[["Name", "Age"], ["Alice", "25"]]')
      --data-file: CSV 또는 JSON 파일에서 데이터 읽기

    예제:
        oa ppt content-add-table --file-path "presentation.pptx" --slide-number 1 --rows 3 --cols 2 --left 1 --top 2 --width 5 --height 3 --data '[["Name", "Age"], ["Alice", "25"], ["Bob", "30"]]' --first-row-header
        oa ppt content-add-table --file-path "presentation.pptx" --slide-number 2 --rows 5 --cols 3 --left 1 --top 1 --width 7 --height 4 --data-file "data.csv" --first-row-header
        oa ppt content-add-table --file-path "presentation.pptx" --slide-number 3 --rows 4 --cols 4 --left 0.5 --top 1.5 --width 9 --height 3.5
    """
    try:
        # 입력 검증
        if data and data_file:
            raise ValueError("--data와 --data-file은 동시에 사용할 수 없습니다")

        if rows < 1 or cols < 1:
            raise ValueError("행과 열은 최소 1 이상이어야 합니다")

        # 파일 경로 정규화 및 존재 확인
        normalized_path = normalize_path(file_path)
        pptx_path = Path(normalized_path).resolve()

        if not pptx_path.exists():
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {pptx_path}")

        # 데이터 로드
        table_data = None
        if data:
            try:
                table_data = json.loads(data)
                if not isinstance(table_data, list):
                    raise ValueError("데이터는 2차원 배열이어야 합니다")
                if table_data and not isinstance(table_data[0], list):
                    raise ValueError("데이터는 2차원 배열이어야 합니다")
            except json.JSONDecodeError as e:
                raise ValueError(f"JSON 데이터 형식이 잘못되었습니다: {str(e)}")

        elif data_file:
            data_file_path = Path(normalize_path(data_file)).resolve()
            if not data_file_path.exists():
                raise FileNotFoundError(f"데이터 파일을 찾을 수 없습니다: {data_file_path}")

            if data_file_path.suffix.lower() == ".csv":
                # CSV 파일 읽기
                with open(data_file_path, "r", encoding="utf-8") as f:
                    reader = csv.reader(f)
                    table_data = list(reader)
            elif data_file_path.suffix.lower() == ".json":
                # JSON 파일 읽기
                with open(data_file_path, "r", encoding="utf-8") as f:
                    table_data = json.load(f)
                    if not isinstance(table_data, list):
                        raise ValueError("JSON 데이터는 2차원 배열이어야 합니다")
                    if table_data and not isinstance(table_data[0], list):
                        raise ValueError("JSON 데이터는 2차원 배열이어야 합니다")
            else:
                raise ValueError("데이터 파일은 .csv 또는 .json 형식이어야 합니다")

        # 데이터 크기 검증
        if table_data:
            data_rows = len(table_data)
            data_cols = max(len(row) for row in table_data) if table_data else 0

            if data_rows > rows:
                raise ValueError(f"데이터 행 수({data_rows})가 표 행 수({rows})보다 큽니다")
            if data_cols > cols:
                raise ValueError(f"데이터 열 수({data_cols})가 표 열 수({cols})보다 큽니다")

        # 프레젠테이션 열기
        prs = Presentation(str(pptx_path))

        # 슬라이드 번호 검증
        slide_idx = validate_slide_number(slide_number, len(prs.slides))
        slide = prs.slides[slide_idx]

        # 표 추가
        table_shape = slide.shapes.add_table(rows, cols, Inches(left), Inches(top), Inches(width), Inches(height))
        table = table_shape.table

        # 데이터 채우기
        if table_data:
            for row_idx, row_data in enumerate(table_data):
                for col_idx, cell_data in enumerate(row_data):
                    if row_idx < rows and col_idx < cols:
                        table.cell(row_idx, col_idx).text = str(cell_data)

        # 헤더 스타일 적용 (첫 행)
        if first_row_header and rows > 0:
            for col_idx in range(cols):
                cell = table.cell(0, col_idx)
                # 헤더 셀 굵게
                for paragraph in cell.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.bold = True

        # 저장
        prs.save(str(pptx_path))

        # 결과 데이터 구성
        result_data = {
            "file": str(pptx_path),
            "slide_number": slide_number,
            "table_size": {"rows": rows, "cols": cols},
            "position": {
                "left": left,
                "top": top,
                "width": width,
                "height": height,
            },
            "first_row_header": first_row_header,
        }

        if table_data:
            result_data["data_filled"] = {
                "rows": len(table_data),
                "cols": max(len(row) for row in table_data) if table_data else 0,
            }

        # 성공 응답
        message = f"슬라이드 {slide_number}에 {rows}×{cols} 표를 추가했습니다"
        if table_data:
            message += f" (데이터 {len(table_data)}행 채움)"

        response = create_success_response(
            data=result_data,
            command="content-add-table",
            message=message,
        )

        # 출력
        if output_format == "json":
            typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
        else:
            typer.echo(f"✅ {message}")
            typer.echo(f"📄 파일: {pptx_path.name}")
            typer.echo(f"📍 슬라이드: {slide_number}")
            typer.echo(f"📊 표 크기: {rows}행 × {cols}열")
            typer.echo(f"📐 위치: {left}in × {top}in")
            typer.echo(f"📏 크기: {width}in × {height}in")
            if first_row_header:
                typer.echo("🎯 첫 행: 헤더로 처리됨")
            if table_data:
                typer.echo(f"💾 데이터: {len(table_data)}행 채워짐")

    except FileNotFoundError as e:
        error_response = create_error_response(e, "content-add-table")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)

    except ValueError as e:
        error_response = create_error_response(e, "content-add-table")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)

    except Exception as e:
        error_response = create_error_response(e, "content-add-table")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ 예기치 않은 오류: {str(e)}", err=True)
        raise typer.Exit(1)


if __name__ == "__main__":
    typer.run(content_add_table)
