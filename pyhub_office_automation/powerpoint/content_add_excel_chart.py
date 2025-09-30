"""
PowerPoint Excel 차트 연동 명령어 (COM 전용)
Excel 워크북의 기존 차트를 PowerPoint 슬라이드에 추가합니다.
"""

import json
import platform
from pathlib import Path
from typing import Optional

import typer

from pyhub_office_automation.version import get_version

from .utils import PowerPointBackend, create_error_response, create_success_response, get_or_open_presentation, normalize_path


def content_add_excel_chart(
    slide_number: int = typer.Option(..., "--slide-number", help="차트를 추가할 슬라이드 번호 (1부터 시작)"),
    excel_file: str = typer.Option(..., "--excel-file", help="Excel 파일 경로"),
    chart_name: Optional[str] = typer.Option(None, "--chart-name", help="차트 이름 (지정하지 않으면 첫 번째 차트)"),
    sheet_name: Optional[str] = typer.Option(None, "--sheet-name", help="시트 이름 (지정하지 않으면 활성 시트)"),
    chart_index: Optional[int] = typer.Option(None, "--chart-index", help="차트 인덱스 (1부터 시작)"),
    left: Optional[float] = typer.Option(None, "--left", help="차트 왼쪽 위치 (인치)"),
    top: Optional[float] = typer.Option(None, "--top", help="차트 상단 위치 (인치)"),
    width: Optional[float] = typer.Option(6.0, "--width", help="차트 너비 (인치, 기본값: 6.0)"),
    height: Optional[float] = typer.Option(4.5, "--height", help="차트 높이 (인치, 기본값: 4.5)"),
    center: bool = typer.Option(False, "--center", help="슬라이드 중앙에 배치 (--left, --top 무시)"),
    link_mode: bool = typer.Option(False, "--link/--embed", help="링크 모드 (기본: 임베드)"),
    file_path: Optional[str] = typer.Option(None, "--file-path", help="PowerPoint 파일 경로"),
    presentation_name: Optional[str] = typer.Option(None, "--presentation-name", help="열려있는 프레젠테이션 이름 (COM 전용)"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 (json/text)"),
):
    """
    Excel 워크북의 기존 차트를 PowerPoint 슬라이드에 추가합니다.

    Windows COM 전용 기능입니다. macOS/Linux에서는 지원되지 않습니다.

    **차트 선택 방법**:
    - --chart-name: 차트 이름으로 찾기 (예: "Chart 1")
    - --chart-index: 차트 인덱스로 찾기 (1부터 시작)
    - 둘 다 지정하지 않으면 첫 번째 차트 사용

    **삽입 모드**:
    - --embed (기본): 독립적인 차트 복사 (Excel 파일과 연결 없음)
    - --link: Excel 차트와 연결 (데이터 업데이트 시 자동 반영)

    **위치 지정**:
    - --center: 슬라이드 중앙에 배치
    - --left, --top: 특정 위치에 배치

    예제:
        # 첫 번째 차트를 중앙에 임베드
        oa ppt content-add-excel-chart --slide-number 2 --excel-file "sales.xlsx" --center

        # 특정 차트를 이름으로 찾아 임베드
        oa ppt content-add-excel-chart --slide-number 3 --excel-file "report.xlsx" --chart-name "Monthly Sales" --left 1 --top 2

        # 특정 시트의 첫 번째 차트를 링크 모드로
        oa ppt content-add-excel-chart --slide-number 4 --excel-file "data.xlsx" --sheet-name "Summary" --link --presentation-name "quarterly.pptx"
    """

    # 1. 플랫폼 체크 (Windows 전용)
    if platform.system() != "Windows":
        result = create_error_response(
            command="content-add-excel-chart",
            error="이 명령어는 Windows에서만 사용 가능합니다 (COM 전용)",
            error_type="PlatformNotSupported",
            details={
                "platform": platform.system(),
                "alternative_suggestions": [
                    "Use content-add-chart to create charts from data",
                    "Export Excel chart as image and use content-add-image",
                ],
            },
        )
        typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
        raise typer.Exit(1)

    try:
        # 2. 입력 검증
        if not center and (left is None or top is None):
            result = create_error_response(
                command="content-add-excel-chart",
                error="--center를 사용하지 않는 경우 --left와 --top을 모두 지정해야 합니다",
                error_type="ValueError",
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

        # Excel 파일 경로 검증
        normalized_excel_path = normalize_path(excel_file)
        excel_path = Path(normalized_excel_path).resolve()

        if not excel_path.exists():
            result = create_error_response(
                command="content-add-excel-chart",
                error=f"Excel 파일을 찾을 수 없습니다: {excel_path}",
                error_type="FileNotFoundError",
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

        # 3. 백엔드는 COM 고정
        selected_backend = PowerPointBackend.COM.value

        # 4. 프레젠테이션 가져오기
        try:
            backend_inst, prs = get_or_open_presentation(
                file_path=file_path,
                presentation_name=presentation_name,
                backend=selected_backend,
            )
        except Exception as e:
            result = create_error_response(
                command="content-add-excel-chart",
                error=f"프레젠테이션을 열 수 없습니다: {str(e)}",
                error_type=type(e).__name__,
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

        # 5. Excel 애플리케이션 열기
        try:
            import win32com.client

            excel_app = win32com.client.Dispatch("Excel.Application")
            excel_app.Visible = False  # 백그라운드 실행
            excel_workbook = excel_app.Workbooks.Open(str(excel_path))

            # 시트 선택
            if sheet_name:
                try:
                    excel_sheet = excel_workbook.Sheets(sheet_name)
                except Exception:
                    excel_workbook.Close(SaveChanges=False)
                    excel_app.Quit()
                    result = create_error_response(
                        command="content-add-excel-chart",
                        error=f"시트를 찾을 수 없습니다: {sheet_name}",
                        error_type="ValueError",
                    )
                    typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
                    raise typer.Exit(1)
            else:
                excel_sheet = excel_workbook.ActiveSheet

            # 차트 찾기
            chart_objects = excel_sheet.ChartObjects()
            if chart_objects.Count == 0:
                excel_workbook.Close(SaveChanges=False)
                excel_app.Quit()
                result = create_error_response(
                    command="content-add-excel-chart",
                    error=f"시트 '{excel_sheet.Name}'에 차트가 없습니다",
                    error_type="ValueError",
                )
                typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
                raise typer.Exit(1)

            # 차트 선택
            excel_chart = None
            if chart_name:
                # 이름으로 찾기
                try:
                    excel_chart = chart_objects(chart_name)
                except Exception:
                    excel_workbook.Close(SaveChanges=False)
                    excel_app.Quit()
                    result = create_error_response(
                        command="content-add-excel-chart",
                        error=f"차트를 찾을 수 없습니다: {chart_name}",
                        error_type="ValueError",
                    )
                    typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
                    raise typer.Exit(1)
            elif chart_index is not None:
                # 인덱스로 찾기
                if chart_index < 1 or chart_index > chart_objects.Count:
                    excel_workbook.Close(SaveChanges=False)
                    excel_app.Quit()
                    result = create_error_response(
                        command="content-add-excel-chart",
                        error=f"차트 인덱스가 범위를 벗어났습니다: {chart_index} (1-{chart_objects.Count})",
                        error_type="ValueError",
                    )
                    typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
                    raise typer.Exit(1)
                excel_chart = chart_objects(chart_index)
            else:
                # 첫 번째 차트
                excel_chart = chart_objects(1)

            chart_name_used = excel_chart.Name

            # 6. 차트 복사
            excel_chart.Copy()

            # 7. PowerPoint 슬라이드 가져오기
            total_slides = prs.Slides.Count
            if slide_number < 1 or slide_number > total_slides:
                excel_workbook.Close(SaveChanges=False)
                excel_app.Quit()
                result = create_error_response(
                    command="content-add-excel-chart",
                    error=f"슬라이드 번호가 범위를 벗어났습니다: {slide_number} (1-{total_slides})",
                    error_type="ValueError",
                )
                typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
                raise typer.Exit(1)

            slide = prs.Slides(slide_number)

            # 8. 위치 계산
            if center:
                slide_width_pt = prs.PageSetup.SlideWidth
                slide_height_pt = prs.PageSetup.SlideHeight
                slide_width_in = slide_width_pt / 72
                slide_height_in = slide_height_pt / 72
                final_left = (slide_width_in - width) / 2
                final_top = (slide_height_in - height) / 2
            else:
                final_left = left
                final_top = top

            # 인치를 포인트로 변환
            left_pt = final_left * 72
            top_pt = final_top * 72
            width_pt = width * 72
            height_pt = height * 72

            # 9. PowerPoint에 붙여넣기
            if link_mode:
                # 링크 모드: ppPasteShape (3)
                shape = slide.Shapes.Paste()
                # 링크 설정 (OLEFormat 사용)
                if hasattr(shape, "LinkFormat"):
                    shape.LinkFormat.SourceFullName = str(excel_path)
                    shape.LinkFormat.AutoUpdate = True
            else:
                # 임베드 모드: 일반 붙여넣기
                shape = slide.Shapes.Paste()

            # 위치 및 크기 설정
            shape.Left = left_pt
            shape.Top = top_pt
            shape.Width = width_pt
            shape.Height = height_pt

            # Excel 정리
            excel_workbook.Close(SaveChanges=False)
            excel_app.Quit()

            # 10. 결과 데이터
            result_data = {
                "backend": "com",
                "slide_number": slide_number,
                "excel_file": str(excel_path),
                "excel_file_name": excel_path.name,
                "sheet_name": excel_sheet.Name,
                "chart_name": chart_name_used,
                "chart_count": chart_objects.Count,
                "position": {
                    "left": round(final_left, 2),
                    "top": round(final_top, 2),
                    "width": width,
                    "height": height,
                },
                "centered": center,
                "link_mode": link_mode,
            }

            message = f"Excel 차트 추가 완료 (COM): 슬라이드 {slide_number}, 차트 '{chart_name_used}'"
            if link_mode:
                message += " (링크 모드)"
            else:
                message += " (임베드 모드)"

        except Exception as e:
            # Excel 정리 (에러 발생 시)
            try:
                if "excel_workbook" in locals():
                    excel_workbook.Close(SaveChanges=False)
                if "excel_app" in locals():
                    excel_app.Quit()
            except Exception:
                pass

            result = create_error_response(
                command="content-add-excel-chart",
                error=f"Excel 차트 추가 실패: {str(e)}",
                error_type=type(e).__name__,
            )
            typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
            raise typer.Exit(1)

        # 11. 성공 응답
        response = create_success_response(
            data=result_data,
            command="content-add-excel-chart",
            message=message,
        )

        # 출력
        if output_format == "json":
            typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
        else:
            typer.echo(f"✅ {message}")
            typer.echo(f"📍 슬라이드: {slide_number}")
            typer.echo(f"📊 Excel 파일: {excel_path.name}")
            typer.echo(f"📄 시트: {excel_sheet.Name}")
            typer.echo(f"📈 차트: {chart_name_used}")
            typer.echo(f"📐 위치: {result_data['position']['left']}in × {result_data['position']['top']}in")
            typer.echo(f"📏 크기: {width}in × {height}in")
            typer.echo(f"🔗 모드: {'링크' if link_mode else '임베드'}")
            typer.echo(f"📊 총 차트 수: {chart_objects.Count}개")

    except typer.Exit:
        raise
    except Exception as e:
        result = create_error_response(
            command="content-add-excel-chart",
            error=str(e),
            error_type=type(e).__name__,
        )
        typer.echo(json.dumps(result, ensure_ascii=False, indent=2))
        raise typer.Exit(1)
    finally:
        # COM 객체 정리는 try-except에서 처리됨
        pass


if __name__ == "__main__":
    typer.run(content_add_excel_chart)
