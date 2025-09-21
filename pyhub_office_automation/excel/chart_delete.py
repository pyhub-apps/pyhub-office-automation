"""
차트 삭제 명령어
워크시트에서 특정 차트를 삭제하는 기능
"""

import json
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


def get_chart_info_before_deletion(chart):
    """삭제 전 차트 정보 수집"""
    try:
        chart_info = {
            "name": chart.name,
            "position": {"left": chart.left, "top": chart.top},
            "dimensions": {"width": chart.width, "height": chart.height},
        }

        # 차트 타입 정보 (가능한 경우)
        try:
            if platform.system() == "Windows":
                chart_type_value = chart.api.ChartType
                # 간단한 차트 타입 매핑
                type_map = {
                    51: "column_clustered",
                    57: "bar_clustered",
                    4: "line",
                    5: "pie",
                    -4120: "doughnut",
                    1: "area",
                    -4169: "scatter",
                    15: "bubble",
                }
                chart_info["chart_type"] = type_map.get(chart_type_value, f"type_{chart_type_value}")
            else:
                chart_info["chart_type"] = "unknown"
        except:
            chart_info["chart_type"] = "unknown"

        # 차트 제목 (가능한 경우)
        try:
            if hasattr(chart, "api") and chart.api.HasTitle:
                chart_info["title"] = chart.api.ChartTitle.Text
        except:
            chart_info["title"] = None

        return chart_info

    except Exception:
        return {"name": getattr(chart, "name", "unknown"), "info_extraction_failed": True}


def chart_delete(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="차트를 삭제할 Excel 파일의 절대 경로"),
    use_active: bool = typer.Option(False, "--use-active", help="현재 활성 워크북 사용"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help='열린 워크북 이름으로 접근 (예: "Sales.xlsx")'),
    sheet: Optional[str] = typer.Option(None, "--sheet", help="차트가 있는 시트 이름 (지정하지 않으면 활성 시트)"),
    chart_name: Optional[str] = typer.Option(None, "--chart-name", help="삭제할 차트의 이름"),
    chart_index: Optional[int] = typer.Option(None, "--chart-index", help="삭제할 차트의 인덱스 (0부터 시작)"),
    all_charts: bool = typer.Option(False, "--all-charts", help="시트의 모든 차트 삭제 (주의: 되돌릴 수 없음)"),
    confirm: bool = typer.Option(False, "--confirm", help="삭제 확인 (--all-charts 사용 시 필수)"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택 (json/text)"),
    visible: bool = typer.Option(False, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부 (기본값: False)"),
    save: bool = typer.Option(True, "--save", help="삭제 후 파일 저장 여부 (기본값: True)"),
):
    """
    워크시트에서 특정 차트를 삭제합니다.

    개별 차트 삭제 또는 시트의 모든 차트를 일괄 삭제할 수 있습니다.
    삭제 작업은 되돌릴 수 없으므로 안전 기능과 함께 신중하게 사용해야 합니다.

    === 워크북 접근 방법 ===
    - --file-path: 파일 경로로 워크북 열기
    - --use-active: 현재 활성 워크북 사용 (권장)
    - --workbook-name: 열린 워크북 이름으로 접근 (예: "Sales.xlsx")

    === 차트 선택 방법 ===
    삭제할 차트를 지정하는 세 가지 방법:

    ▶ 개별 차트 선택:
      • --chart-name "Chart1": 차트 이름으로 선택
      • --chart-index 0: 차트 인덱스로 선택 (0부터 시작)
      • chart-list 명령으로 차트 이름/인덱스 확인 가능

    ▶ 전체 차트 삭제:
      • --all-charts: 시트의 모든 차트 삭제
      • --confirm 플래그 필수 (안전 장치)
      • 삭제 전 각 차트 정보 수집 후 진행

    === 안전 기능 및 보안 ===
    ▶ 삭제 전 정보 백업:
      • 차트 이름, 위치, 크기, 유형 저장
      • 차트 제목 및 설정 정보 보관
      • JSON 응답에 삭제된 차트 정보 포함

    ▶ 안전 장치:
      • --all-charts 사용 시 --confirm 플래그 의무
      • 삭제될 차트 수 및 정보 사전 표시
      • 상세 출력으로 삭제 결과 확인 가능

    === 차트 삭제 시나리오 ===
    ▶ 개별 차트 삭제:
      • 불필요한 차트 제거
      • 오래된 차트 정리
      • 잘못 생성된 차트 수정

    ▶ 전체 차트 삭제:
      • 시트 리셋 및 재설계
      • 새로운 대시보드 구성
      • 템플릿 초기화

    === 실제 활용 예제 ===

    # 1. 첫 번째 차트 삭제
    oa excel chart-delete --use-active --chart-index 0

    # 2. 이름으로 특정 차트 삭제
    oa excel chart-delete --file-path "report.xlsx" --chart-name "SalesChart"

    # 3. 특정 시트의 차트 삭제
    oa excel chart-delete --workbook-name "Dashboard.xlsx" --sheet "Charts" --chart-index 1

    # 4. 전체 차트 삭제 (안전 확인 후)
    oa excel chart-delete --use-active --sheet "Dashboard" --all-charts --confirm

    # 5. 삭제 결과 텍스트 형식으로 확인
    oa excel chart-delete --workbook-name "Old_Report.xlsx" --all-charts --confirm --format text

    # 6. 저장 없이 미리보기
    oa excel chart-delete --use-active --chart-index 0 --save false

    === 삭제 전 체크리스트 ===
    • chart-list 명령으로 삭제 대상 차트 확인
    • 중요한 차트는 chart-export로 백업 고려
    • --all-charts 사용 시 반드시 --confirm 플래그 포함
    • 삭제 후 되돌릴 수 없음을 명심

    === 예외 상황 처리 ===
    • 차트가 없는 시트: 안전하게 오류 메시지 반환
    • 존재하지 않는 차트 이름/인덱스: 사용 가능한 차트 목록 제공
    • 개별 차트 삭제 실패: 이어서 다른 차트 삭제 진행

    === 주의사항 ===
    ⚠️ **삭제된 차트는 복구할 수 없습니다**
    ⚠️ 중요한 차트는 사전에 백업 및 내보내기 권장
    ⚠️ --all-charts 옵션은 전체 대시보드에 영향을 줄 수 있음
    """
    # 입력 값 검증
    if output_format not in ["json", "text"]:
        raise ValueError(f"잘못된 출력 형식: {output_format}. 사용 가능한 형식: json, text")

    book = None

    try:
        # 워크북 연결
        book = get_or_open_workbook(file_path=file_path, workbook_name=workbook_name, use_active=use_active, visible=visible)

        # 시트 가져오기
        target_sheet = get_sheet(book, sheet)

        # 시트에 차트가 있는지 확인
        if len(target_sheet.charts) == 0:
            raise ValueError(f"시트 '{target_sheet.name}'에 삭제할 차트가 없습니다")

        deleted_charts = []
        deletion_summary = {
            "sheet": target_sheet.name,
            "workbook": book.name,
            "total_charts_before": len(target_sheet.charts),
            "deleted_charts": [],
            "total_deleted": 0,
            "remaining_charts": 0,
        }

        if all_charts:
            # 모든 차트 삭제
            if not confirm:
                raise ValueError("모든 차트를 삭제하려면 --confirm 플래그가 필요합니다")

            # 차트 정보를 먼저 수집 (삭제 전)
            charts_to_delete = []
            for i, chart in enumerate(target_sheet.charts):
                chart_info = get_chart_info_before_deletion(chart)
                chart_info["index"] = i
                charts_to_delete.append((chart, chart_info))

            # 모든 차트 삭제 (역순으로 삭제하여 인덱스 문제 방지)
            for chart, chart_info in reversed(charts_to_delete):
                try:
                    chart.delete()
                    deletion_summary["deleted_charts"].append(chart_info)
                    deletion_summary["total_deleted"] += 1
                except Exception as e:
                    # 개별 차트 삭제 실패시에도 계속 진행
                    chart_info["deletion_error"] = str(e)
                    deletion_summary["deleted_charts"].append(chart_info)

        else:
            # 개별 차트 삭제
            chart = find_chart_by_name_or_index(target_sheet, chart_name, chart_index)

            # 삭제 전 차트 정보 수집
            chart_info = get_chart_info_before_deletion(chart)

            # 차트 삭제
            chart.delete()

            deletion_summary["deleted_charts"].append(chart_info)
            deletion_summary["total_deleted"] = 1

        # 삭제 후 남은 차트 수 계산
        deletion_summary["remaining_charts"] = len(target_sheet.charts)

        # 파일 저장
        if save and file_path:
            book.save()
            deletion_summary["file_saved"] = True

        # 응답 생성
        if deletion_summary["total_deleted"] > 0:
            if all_charts:
                message = f"시트 '{target_sheet.name}'에서 {deletion_summary['total_deleted']}개의 차트를 삭제했습니다"
            else:
                chart_name_deleted = deletion_summary["deleted_charts"][0]["name"]
                message = f"차트 '{chart_name_deleted}'을 삭제했습니다"
        else:
            message = "삭제된 차트가 없습니다"

        response = create_success_response(data=deletion_summary, command="chart-delete", message=message)

        if output_format == "json":
            print(json.dumps(response, ensure_ascii=False, indent=2))
        else:
            # 텍스트 형식 출력
            print(f"=== 차트 삭제 결과 ===")
            print(f"시트: {target_sheet.name}")
            print(f"삭제 전 차트 수: {deletion_summary['total_charts_before']}")
            print(f"삭제된 차트 수: {deletion_summary['total_deleted']}")
            print(f"남은 차트 수: {deletion_summary['remaining_charts']}")
            print()

            if deletion_summary["deleted_charts"]:
                print("🗑️ 삭제된 차트:")
                for chart_info in deletion_summary["deleted_charts"]:
                    print(f"   📊 {chart_info['name']}")
                    if chart_info.get("title"):
                        print(f"      제목: {chart_info['title']}")
                    if chart_info.get("chart_type"):
                        print(f"      유형: {chart_info['chart_type']}")
                    if chart_info.get("deletion_error"):
                        print(f"      ❌ 삭제 오류: {chart_info['deletion_error']}")
                    else:
                        print(f"      ✅ 삭제 완료")
                print()

            if save and file_path:
                print("💾 파일이 저장되었습니다.")

            if all_charts and deletion_summary["total_deleted"] > 0:
                print("⚠️ 모든 차트가 삭제되었습니다. 이 작업은 되돌릴 수 없습니다.")

    except Exception as e:
        error_response = create_error_response(e, "chart-delete")
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
    chart_delete()
