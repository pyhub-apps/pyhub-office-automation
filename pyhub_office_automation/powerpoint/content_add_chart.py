"""
PowerPoint 차트 추가 명령어
Excel 데이터 또는 CSV 파일로부터 차트를 생성합니다.
"""

import json
from pathlib import Path
from typing import Optional

import typer
from pptx import Presentation
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches

from pyhub_office_automation.version import get_version

from .utils import (
    ChartType,
    create_chart_data,
    create_error_response,
    create_success_response,
    load_data_from_csv,
    load_data_from_excel,
    normalize_path,
    parse_excel_range,
    validate_slide_number,
)

# 차트 타입 매핑
CHART_TYPE_MAP = {
    ChartType.COLUMN.value: XL_CHART_TYPE.COLUMN_CLUSTERED,
    ChartType.BAR.value: XL_CHART_TYPE.BAR_CLUSTERED,
    ChartType.LINE.value: XL_CHART_TYPE.LINE,
    ChartType.PIE.value: XL_CHART_TYPE.PIE,
    ChartType.AREA.value: XL_CHART_TYPE.AREA,
    ChartType.SCATTER.value: XL_CHART_TYPE.XY_SCATTER,
    ChartType.DOUGHNUT.value: XL_CHART_TYPE.DOUGHNUT,
}


def content_add_chart(
    file_path: str = typer.Option(..., "--file-path", help="PowerPoint 파일 경로"),
    slide_number: int = typer.Option(..., "--slide-number", help="차트를 추가할 슬라이드 번호 (1부터 시작)"),
    chart_type: str = typer.Option(..., "--chart-type", help="차트 타입 (column/bar/line/pie/area/scatter/doughnut)"),
    csv_data: Optional[str] = typer.Option(None, "--csv-data", help="CSV 파일 경로"),
    excel_data: Optional[str] = typer.Option(None, "--excel-data", help="Excel 데이터 참조 (예: data.xlsx!A1:C10)"),
    left: Optional[float] = typer.Option(None, "--left", help="차트 왼쪽 위치 (인치)"),
    top: Optional[float] = typer.Option(None, "--top", help="차트 상단 위치 (인치)"),
    width: Optional[float] = typer.Option(6.0, "--width", help="차트 너비 (인치, 기본값: 6.0)"),
    height: Optional[float] = typer.Option(4.5, "--height", help="차트 높이 (인치, 기본값: 4.5)"),
    title: Optional[str] = typer.Option(None, "--title", help="차트 제목"),
    center: bool = typer.Option(False, "--center", help="슬라이드 중앙에 배치 (--left, --top 무시)"),
    show_legend: bool = typer.Option(True, "--show-legend/--no-legend", help="범례 표시 여부 (기본값: 표시)"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 (json/text)"),
):
    """
    PowerPoint 슬라이드에 데이터 기반 차트를 추가합니다.

    데이터 소스 (둘 중 하나만 지정):
      --csv-data: CSV 파일 경로
      --excel-data: Excel 참조 (예: "data.xlsx!A1:C10" 또는 "data.xlsx!Sheet1!A1:C10")

    차트 타입:
      column, bar, line, pie, area, scatter, doughnut

    위치 지정:
      --center: 슬라이드 중앙에 배치
      --left, --top: 특정 위치에 배치

    예제:
        oa ppt content-add-chart --file-path "presentation.pptx" --slide-number 2 --chart-type column --csv-data "sales.csv" --center --title "판매 현황"
        oa ppt content-add-chart --file-path "presentation.pptx" --slide-number 3 --chart-type pie --excel-data "data.xlsx!A1:C10" --left 1 --top 2
    """
    try:
        # 입력 검증
        if not csv_data and not excel_data:
            raise ValueError("--csv-data 또는 --excel-data 중 하나는 반드시 지정해야 합니다")

        if csv_data and excel_data:
            raise ValueError("--csv-data와 --excel-data는 동시에 사용할 수 없습니다")

        if not center and (left is None or top is None):
            raise ValueError("--center를 사용하지 않는 경우 --left와 --top을 모두 지정해야 합니다")

        # 차트 타입 검증
        chart_type_lower = chart_type.lower()
        if chart_type_lower not in CHART_TYPE_MAP:
            available_types = ", ".join(CHART_TYPE_MAP.keys())
            raise ValueError(f"지원하지 않는 차트 타입: {chart_type}\n사용 가능: {available_types}")

        # 파일 경로 정규화 및 존재 확인
        normalized_pptx_path = normalize_path(file_path)
        pptx_path = Path(normalized_pptx_path).resolve()

        if not pptx_path.exists():
            raise FileNotFoundError(f"PowerPoint 파일을 찾을 수 없습니다: {pptx_path}")

        # 데이터 로드
        df = None
        data_source = None

        if csv_data:
            df = load_data_from_csv(csv_data)
            data_source = str(Path(csv_data).name)
        else:
            # Excel 데이터 파싱
            excel_ref = parse_excel_range(excel_data)
            df = load_data_from_excel(
                file_path=excel_ref["file_path"], sheet_name=excel_ref["sheet"], range_addr=excel_ref["range"]
            )
            data_source = excel_data

        # 데이터 검증
        if df is None or df.empty:
            raise ValueError("데이터가 비어있습니다")

        if len(df.columns) < 2:
            raise ValueError(f"차트를 생성하려면 최소 2개의 열이 필요합니다 (현재: {len(df.columns)}개)")

        # ChartData 생성
        chart_data = create_chart_data(df, chart_type_lower)

        # 프레젠테이션 열기
        prs = Presentation(str(pptx_path))

        # 슬라이드 번호 검증
        slide_idx = validate_slide_number(slide_number, len(prs.slides))
        slide = prs.slides[slide_idx]

        # 위치 계산
        if center:
            # 슬라이드 크기 가져오기 (EMU 단위)
            slide_width = prs.slide_width
            slide_height = prs.slide_height

            # 인치 단위로 변환
            slide_width_in = slide_width / 914400  # 1 inch = 914400 EMU
            slide_height_in = slide_height / 914400

            # 중앙 배치 위치 계산
            final_left = (slide_width_in - width) / 2
            final_top = (slide_height_in - height) / 2
        else:
            final_left = left
            final_top = top

        # 차트 추가
        chart_type_const = CHART_TYPE_MAP[chart_type_lower]
        graphic_frame = slide.shapes.add_chart(
            chart_type_const, Inches(final_left), Inches(final_top), Inches(width), Inches(height), chart_data
        )

        chart = graphic_frame.chart

        # 차트 제목 설정
        if title:
            chart.has_title = True
            chart.chart_title.text_frame.text = title

        # 범례 설정
        chart.has_legend = show_legend

        # 저장
        prs.save(str(pptx_path))

        # 결과 데이터 구성
        result_data = {
            "file": str(pptx_path),
            "slide_number": slide_number,
            "chart_type": chart_type_lower,
            "data_source": data_source,
            "data_shape": {"rows": len(df), "columns": len(df.columns)},
            "series_count": len(df.columns) - 1,  # 첫 번째 열은 카테고리
            "position": {
                "left": round(final_left, 2),
                "top": round(final_top, 2),
                "width": width,
                "height": height,
            },
            "centered": center,
            "has_title": title is not None,
            "has_legend": show_legend,
        }

        if title:
            result_data["title"] = title

        # 성공 응답
        message = f"슬라이드 {slide_number}에 {chart_type_lower} 차트를 추가했습니다"
        if title:
            message += f" (제목: {title})"

        response = create_success_response(
            data=result_data,
            command="content-add-chart",
            message=message,
        )

        # 출력
        if output_format == "json":
            typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
        else:
            typer.echo(f"✅ {message}")
            typer.echo(f"📄 파일: {pptx_path.name}")
            typer.echo(f"📍 슬라이드: {slide_number}")
            typer.echo(f"📊 차트 타입: {chart_type_lower}")
            typer.echo(f"📈 데이터 소스: {data_source}")
            typer.echo(f"📏 데이터 크기: {result_data['data_shape']['rows']}행 × {result_data['data_shape']['columns']}열")
            typer.echo(f"📐 위치: {result_data['position']['left']}in × {result_data['position']['top']}in")
            typer.echo(f"📏 크기: {width}in × {height}in")
            if title:
                typer.echo(f"🏷️ 제목: {title}")
            typer.echo(f"📊 시리즈 개수: {result_data['series_count']}")
            typer.echo(f"📖 범례: {'표시' if show_legend else '숨김'}")

    except FileNotFoundError as e:
        error_response = create_error_response(e, "content-add-chart")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)

    except ValueError as e:
        error_response = create_error_response(e, "content-add-chart")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)

    except Exception as e:
        error_response = create_error_response(e, "content-add-chart")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ 예기치 않은 오류: {str(e)}", err=True)
        raise typer.Exit(1)


if __name__ == "__main__":
    typer.run(content_add_chart)
