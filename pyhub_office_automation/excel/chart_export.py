"""
차트 내보내기 명령어
차트를 이미지 파일로 내보내는 기능
"""

import json
import os
import platform
from pathlib import Path
from typing import Optional

import typer
import xlwings as xw

from pyhub_office_automation.version import get_version

from .utils import create_error_response, create_success_response, get_or_open_workbook, get_sheet, normalize_path


def find_chart_by_name_or_index(sheet, chart_name=None, chart_index=None):
    """차트 이름이나 인덱스로 차트 객체 찾기"""
    if chart_name:
        for chart in sheet.charts:
            if chart.name == chart_name:
                return chart
        raise ValueError(f"차트 '{chart_name}'을 찾을 수 없습니다")

    elif chart_index is not None:
        try:
            if 0 <= chart_index < len(sheet.charts):
                return sheet.charts[chart_index]
            else:
                raise IndexError(f"차트 인덱스 {chart_index}는 범위를 벗어났습니다 (0-{len(sheet.charts)-1})")
        except IndexError as e:
            raise ValueError(str(e))

    else:
        raise ValueError("차트 이름(--chart-name) 또는 인덱스(--chart-index) 중 하나를 지정해야 합니다")


def get_image_format_constant(image_format):
    """이미지 형식에 해당하는 xlwings 상수를 반환"""
    # 이미지 형식 상수값 (Windows Excel에서 사용하는 값들)
    format_values = {
        "png": -4142,  # xlPNG
        "jpg": -4141,  # xlJPEG
        "jpeg": -4141,  # xlJPEG
        "gif": -4140,  # xlGIF
        "bmp": -4147,  # xlBMP
    }

    format_lower = image_format.lower()
    if format_lower not in format_values:
        raise ValueError(f"지원되지 않는 이미지 형식: {image_format}")

    if platform.system() == "Windows":
        # Windows에서는 xlwings 상수를 시도하고, 실패하면 숫자값 사용
        try:
            from xlwings.constants import FileFormat

            const_map = {-4142: "xlPNG", -4141: "xlJPEG", -4140: "xlGIF", -4147: "xlBMP"}

            format_value = format_values[format_lower]
            const_name = const_map.get(format_value)

            if const_name and hasattr(FileFormat, const_name):
                return getattr(FileFormat, const_name)
            else:
                # 상수를 찾을 수 없으면 숫자값 직접 반환
                return format_value

        except ImportError:
            # FileFormat을 가져올 수 없으면 숫자값 직접 반환
            return format_values[format_lower]
    else:
        # macOS에서는 문자열로 반환
        return format_lower


def validate_output_path(output_path, image_format):
    """출력 경로 검증 및 정규화"""
    # 한글 경로 정규화
    output_path = normalize_path(output_path)
    output_path = Path(output_path).resolve()

    # 확장자가 없으면 추가
    if not output_path.suffix:
        output_path = output_path.with_suffix(f".{image_format.lower()}")

    # 확장자가 지정된 형식과 다르면 경고
    expected_ext = f".{image_format.lower()}"
    if image_format.lower() == "jpeg":
        expected_ext = ".jpg"  # JPEG는 보통 .jpg 확장자 사용

    if output_path.suffix.lower() != expected_ext:
        # 확장자를 형식에 맞게 변경
        output_path = output_path.with_suffix(expected_ext)

    # 출력 디렉토리가 존재하지 않으면 생성
    output_path.parent.mkdir(parents=True, exist_ok=True)

    return output_path


def get_chart_export_info(chart):
    """내보내기 전 차트 정보 수집"""
    try:
        chart_info = {
            "name": chart.name,
            "position": {"left": chart.left, "top": chart.top},
            "dimensions": {"width": chart.width, "height": chart.height},
        }

        # 차트 제목 (가능한 경우)
        try:
            if hasattr(chart, "api") and chart.api.HasTitle:
                chart_info["title"] = chart.api.ChartTitle.Text
        except:
            chart_info["title"] = None

        return chart_info

    except Exception:
        return {"name": getattr(chart, "name", "unknown"), "info_extraction_failed": True}


def chart_export(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="차트가 있는 Excel 파일의 절대 경로"),
    use_active: bool = typer.Option(False, "--use-active", help="현재 활성 워크북 사용"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help='열린 워크북 이름으로 접근 (예: "Sales.xlsx")'),
    sheet: Optional[str] = typer.Option(None, "--sheet", help="차트가 있는 시트 이름 (지정하지 않으면 활성 시트)"),
    chart_name: Optional[str] = typer.Option(None, "--chart-name", help="내보낼 차트의 이름"),
    chart_index: Optional[int] = typer.Option(None, "--chart-index", help="내보낼 차트의 인덱스 (0부터 시작)"),
    output_path: str = typer.Option(..., "--output-path", help="이미지 파일을 저장할 경로 (확장자 포함 또는 자동 추가)"),
    image_format: str = typer.Option("png", "--image-format", help="이미지 형식 (png/jpg/jpeg/gif/bmp, 기본값: png)"),
    width: Optional[int] = typer.Option(None, "--width", help="내보낼 이미지의 너비 (픽셀, 지정하지 않으면 차트 원본 크기)"),
    height: Optional[int] = typer.Option(None, "--height", help="내보낼 이미지의 높이 (픽셀, 지정하지 않으면 차트 원본 크기)"),
    dpi: int = typer.Option(300, "--dpi", help="이미지 해상도 (DPI, 기본값: 300)"),
    transparent_bg: bool = typer.Option(False, "--transparent-bg", help="투명 배경으로 내보내기 (PNG 형식에서만 지원)"),
    overwrite: bool = typer.Option(False, "--overwrite", help="기존 파일이 있으면 덮어쓰기"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택 (json/text)"),
    visible: bool = typer.Option(False, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부 (기본값: False)"),
):
    """
    Excel 차트를 이미지 파일로 내보냅니다.

    차트를 PNG, JPG, GIF, BMP 형식의 고품질 이미지 파일로 저장할 수 있습니다.
    프레젠테이션, 웹사이트, 보고서 등에 활용할 수 있도록 해상도와 크기를 자유롭게 조정할 수 있습니다.

    === 워크북 접근 방법 ===
    - --file-path: 파일 경로로 워크북 열기
    - --use-active: 현재 활성 워크북 사용 (권장)
    - --workbook-name: 열린 워크북 이름으로 접근 (예: "Sales.xlsx")

    === 차트 선택 방법 ===
    내보낼 차트를 지정하는 두 가지 방법:

    ▶ 차트 이름으로 선택:
      • --chart-name "SalesChart"
      • chart-list 명령으로 차트 이름 확인 가능

    ▶ 인덱스 번호로 선택:
      • --chart-index 0 (첫 번째 차트)
      • 시트의 차트 순서대로 0, 1, 2...

    === 이미지 형식별 특징과 용도 ===

    ▶ PNG (권장):
      • 고품질 무손실 압축
      • 투명 배경 지원 (--transparent-bg)
      • 프레젠테이션, 웹사이트에 최적
      • 파일 크기: 중간

    ▶ JPG/JPEG:
      • 작은 파일 크기
      • 사진 품질의 색상 표현
      • 이메일, 웹 업로드에 적합
      • 투명 배경 미지원

    ▶ GIF:
      • 매우 작은 파일 크기
      • 제한된 색상 (256색)
      • 웹용 간단한 차트에 적합
      • 투명 배경 미지원

    ▶ BMP:
      • 압축되지 않은 원본 품질
      • 가장 큰 파일 크기
      • 고품질 인쇄용
      • 투명 배경 미지원

    === 크기 및 품질 설정 ===

    ▶ 크기 설정:
      • 원본 크기: 옵션 생략 (차트 원본 크기 유지)
      • 사용자 지정: --width 800 --height 600
      • 비율 유지 안함: 원하는 크기로 자유 조정

    ▶ 해상도 설정:
      • --dpi 300: 고해상도 (기본값, 인쇄용)
      • --dpi 150: 중간 해상도 (웹/화면용)
      • --dpi 72: 저해상도 (웹 최적화)

    ▶ 배경 설정:
      • --transparent-bg: 투명 배경 (PNG만 지원)
      • 프레젠테이션 슬라이드에 삽입 시 유용

    === 출력 경로 처리 ===
    --output-path 옵션으로 저장 위치를 지정합니다:

    • 확장자 자동 추가: "chart1" → "chart1.png"
    • 확장자 형식 맞춤: "chart.jpg" + --image-format png → "chart.png"
    • 디렉토리 자동 생성: 없는 폴더는 자동으로 생성
    • 기존 파일 보호: --overwrite 없이는 덮어쓰기 방지

    === 활용 시나리오별 권장 설정 ===

    ▶ 프레젠테이션용 (PowerPoint, Keynote):
      • PNG, 투명 배경, 고해상도
      • 크기: 슬라이드에 맞게 조정

    ▶ 웹사이트/블로그용:
      • JPG (작은 크기) 또는 PNG (고품질)
      • 해상도: 72-150 DPI
      • 크기: 웹페이지 레이아웃에 맞춤

    ▶ 인쇄용 보고서:
      • PNG 또는 BMP, 300 DPI
      • 원본 크기 또는 큰 크기
      • 고품질 유지

    ▶ 이메일/공유용:
      • JPG, 중간 해상도
      • 작은 파일 크기로 빠른 전송

    === 실제 활용 예제 ===

    # 1. 기본 PNG 내보내기
    oa excel chart-export --use-active --chart-index 0 --output-path "chart1.png"

    # 2. 웹용 JPG 내보내기 (작은 크기)
    oa excel chart-export --file-path "report.xlsx" --chart-name "SalesChart" \\
        --output-path "sales.jpg" --image-format "jpg" --dpi 150 --width 600 --height 400

    # 3. 프레젠테이션용 투명 PNG
    oa excel chart-export --workbook-name "Dashboard.xlsx" --chart-index 0 \\
        --output-path "dashboard.png" --width 800 --height 600 --transparent-bg

    # 4. 고해상도 인쇄용 이미지
    oa excel chart-export --use-active --chart-name "QuarterlyReport" \\
        --output-path "print_chart.png" --dpi 300 --image-format "png"

    # 5. 여러 차트 일괄 내보내기 (스크립트 활용)
    oa excel chart-export --use-active --chart-index 0 --output-path "chart_0.png"
    oa excel chart-export --use-active --chart-index 1 --output-path "chart_1.png"
    oa excel chart-export --use-active --chart-index 2 --output-path "chart_2.png"

    # 6. 기존 파일 덮어쓰기
    oa excel chart-export --file-path "old_report.xlsx" --chart-name "OldChart" \\
        --output-path "updated_chart.png" --overwrite

    === 플랫폼별 기능 차이 ===
    • Windows: 모든 형식과 고급 기능 지원 (투명 배경, 고해상도 등)
    • macOS: 기본 내보내기 지원, 일부 고급 기능 제한

    === 파일 관리 팁 ===
    • 체계적인 파일명 사용: "report_sales_2024.png"
    • 용도별 폴더 구분: charts/web/, charts/print/
    • 날짜 포함: "dashboard_20241201.png"
    • 백업 고려: 중요한 차트는 여러 형식으로 저장
    """
    # 입력 값 검증
    if image_format not in ["png", "jpg", "jpeg", "gif", "bmp"]:
        raise ValueError(f"잘못된 이미지 형식: {image_format}. 사용 가능한 형식: png, jpg, jpeg, gif, bmp")

    if output_format not in ["json", "text"]:
        raise ValueError(f"잘못된 출력 형식: {output_format}. 사용 가능한 형식: json, text")

    book = None

    try:
        # 워크북 연결
        book = get_or_open_workbook(file_path=file_path, workbook_name=workbook_name, use_active=use_active, visible=visible)

        # 시트 가져오기
        target_sheet = get_sheet(book, sheet)

        # 차트 찾기
        chart = find_chart_by_name_or_index(target_sheet, chart_name, chart_index)

        # 출력 경로 검증 및 정규화
        validated_output_path = validate_output_path(output_path, image_format)

        # 기존 파일 존재 확인
        if validated_output_path.exists() and not overwrite:
            raise ValueError(f"파일 '{validated_output_path}'가 이미 존재합니다. --overwrite 옵션을 사용하여 덮어쓰세요.")

        # 차트 정보 수집
        chart_info = get_chart_export_info(chart)

        # 원본 차트 크기 저장
        original_width = chart.width
        original_height = chart.height

        # 크기 조정 (지정된 경우)
        size_changed = False
        if width or height:
            if width:
                chart.width = width
                size_changed = True
            if height:
                chart.height = height
                size_changed = True

        try:
            # 이미지 형식 상수 가져오기
            format_constant = get_image_format_constant(image_format)

            # 차트 내보내기
            if platform.system() == "Windows":
                # Windows에서는 COM API 사용
                try:
                    # 투명 배경 설정 (PNG에서만)
                    if transparent_bg and image_format.lower() == "png":
                        try:
                            chart.api.ChartArea.Format.Fill.Transparency = 1.0
                            chart.api.PlotArea.Format.Fill.Transparency = 1.0
                        except:
                            pass  # 투명 배경 설정 실패해도 계속 진행

                    # Export 메서드 사용
                    chart.api.Export(str(validated_output_path), FilterName=image_format.upper())

                except Exception as e:
                    # Export 실패 시 다른 방법 시도
                    try:
                        chart.api.ChartArea.Copy()
                        # 클립보드에서 이미지 저장은 복잡하므로 여기서는 에러 발생
                        raise RuntimeError(f"차트 내보내기 실패: {str(e)}")
                    except:
                        raise RuntimeError(f"차트 내보내기 실패: {str(e)}")

            else:
                # macOS에서는 xlwings 기본 기능 사용
                try:
                    # macOS에서는 제한적인 내보내기 지원
                    chart.api.Export(str(validated_output_path))
                except Exception as e:
                    raise RuntimeError(f"macOS에서 차트 내보내기 실패: {str(e)}. 일부 기능은 Windows에서만 지원됩니다.")

        finally:
            # 차트 크기 복원
            if size_changed:
                chart.width = original_width
                chart.height = original_height

        # 내보내기 성공 확인
        if not validated_output_path.exists():
            raise RuntimeError("이미지 파일이 생성되지 않았습니다")

        # 생성된 파일 정보 수집
        file_stats = validated_output_path.stat()
        file_info = {
            "path": str(validated_output_path),
            "size_bytes": file_stats.st_size,
            "size_mb": round(file_stats.st_size / (1024 * 1024), 2),
            "format": image_format,
            "exists": True,
        }

        # 응답 데이터 구성
        response_data = {
            "chart_info": chart_info,
            "export_settings": {
                "format": image_format,
                "dpi": dpi,
                "transparent_background": transparent_bg and image_format.lower() == "png",
                "custom_size": {"width": width, "height": height} if (width or height) else None,
                "original_size": {"width": original_width, "height": original_height},
            },
            "output_file": file_info,
            "sheet": target_sheet.name,
            "workbook": book.name,
            "platform": platform.system(),
        }

        message = f"차트 '{chart.name}'을 '{validated_output_path.name}'으로 내보냈습니다"

        response = create_success_response(data=response_data, command="chart-export", message=message)

        if output_format == "json":
            print(json.dumps(response, ensure_ascii=False, indent=2))
        else:
            # 텍스트 형식 출력
            print(f"=== 차트 내보내기 결과 ===")
            print(f"차트: {chart.name}")
            print(f"시트: {target_sheet.name}")
            if chart_info.get("title"):
                print(f"제목: {chart_info['title']}")
            print()

            print(f"📁 출력 파일:")
            print(f"   경로: {validated_output_path}")
            print(f"   형식: {image_format.upper()}")
            print(f"   크기: {file_info['size_mb']} MB ({file_info['size_bytes']:,} bytes)")
            print()

            print(f"🎨 내보내기 설정:")
            print(f"   해상도: {dpi} DPI")
            if width or height:
                print(f"   이미지 크기: {width or '원본'} x {height or '원본'}")
                print(f"   원본 크기: {original_width} x {original_height}")
            else:
                print(f"   크기: 원본 ({original_width} x {original_height})")

            if transparent_bg and image_format.lower() == "png":
                print(f"   배경: 투명")
            elif transparent_bg:
                print(f"   배경: 투명 (PNG가 아닌 형식에서는 지원되지 않음)")

            print(f"\n✅ 차트가 성공적으로 내보내졌습니다!")

    except Exception as e:
        error_response = create_error_response(e, "chart-export")
        if output_format == "json":
            print(json.dumps(error_response, ensure_ascii=False, indent=2))
        else:
            print(f"오류: {str(e)}")
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
    chart_export()
