"""
도형 목록 조회 명령어
xlwings를 활용한 Excel 도형 정보 수집 기능
"""

import json
import platform
from typing import Optional

import typer
import xlwings as xw

from pyhub_office_automation.version import get_version

from .utils import (
    ExecutionTimer,
    create_error_response,
    create_success_response,
    get_or_open_workbook,
    get_shapes_info,
    get_sheet,
    normalize_path,
)


def shape_list(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="도형을 조회할 Excel 파일의 절대 경로"),
    use_active: bool = typer.Option(False, "--use-active", help="현재 활성 워크북 사용"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help='열린 워크북 이름으로 접근 (예: "Sales.xlsx")'),
    sheet: Optional[str] = typer.Option(None, "--sheet", help="도형을 조회할 시트 이름 (지정하지 않으면 활성 시트)"),
    detailed: bool = typer.Option(False, "--detailed", help="상세 정보 포함 (색상, 투명도, 텍스트 등)"),
    include_text: bool = typer.Option(False, "--include-text", help="텍스트 내용 포함 (Windows 전용)"),
    filter_type: Optional[str] = typer.Option(None, "--filter-type", help="특정 도형 타입만 필터링 (예: rectangle, oval)"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택 (json/text)"),
    visible: bool = typer.Option(False, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부 (기본값: False)"),
):
    """
    Excel 시트의 모든 도형 정보를 조회합니다.

    도형의 기본 정보(이름, 위치, 크기)부터 상세 정보(색상, 스타일, 텍스트)까지
    조회할 수 있으며, 대시보드 분석 및 도형 관리에 유용합니다.

    === 워크북 접근 방법 ===
    - --file-path: 파일 경로로 워크북 열기
    - --use-active: 현재 활성 워크북 사용
    - --workbook-name: 열린 워크북 이름으로 접근 (예: "Sales.xlsx")

    === 조회 옵션 ===
    • 기본 조회: 도형 이름, 타입, 위치, 크기
    • --detailed: 색상, 투명도, 가시성 등 상세 정보 포함
    • --include-text: 텍스트 박스나 도형 내 텍스트 내용 포함 (Windows)
    • --filter-type: 특정 타입의 도형만 필터링

    === 기본 정보 항목 ===
    • name: 도형 이름
    • type: 도형 타입 (Windows에서 숫자 코드)
    • position: {left, top} 위치 정보
    • size: {width, height} 크기 정보

    === 상세 정보 항목 (--detailed) ===
    • fill_color: 채우기 색상 (HEX 형식)
    • transparency: 투명도 (0-100)
    • visible: 가시성 여부
    • has_text: 텍스트 포함 여부

    === 사용 시나리오 ===

    # 1. 활성 시트의 모든 도형 기본 정보 조회
    oa excel shape-list --use-active

    # 2. 특정 시트의 상세한 도형 정보 조회
    oa excel shape-list --file-path "dashboard.xlsx" --sheet "Dashboard" --detailed

    # 3. 특정 타입의 도형만 필터링
    oa excel shape-list --use-active --filter-type "rectangle" --detailed

    # 4. 텍스트 내용까지 포함한 전체 조회 (Windows)
    oa excel shape-list --use-active --detailed --include-text

    # 5. 대시보드 분석을 위한 전체 정보 수집
    oa excel shape-list --workbook-name "Report.xlsx" --sheet "Dashboard" \\
        --detailed --include-text --format json

    === 출력 예제 ===
    ```json
    {
      "success": true,
      "data": {
        "shapes": [
          {
            "name": "TitleArea",
            "type": 1,
            "position": {"left": 70, "top": 70},
            "size": {"width": 760, "height": 80},
            "fill_color": "#1D2433",
            "transparency": 0,
            "visible": true,
            "has_text": false
          }
        ],
        "total_shapes": 1,
        "sheet": "Dashboard"
      }
    }
    ```

    === 대시보드 관리 활용 ===
    • 뉴모피즘 도형 현황 파악
    • 레이아웃 구조 분석
    • 색상 일관성 검토
    • 도형 이름 체계 확인
    • 위치 좌표 수집 (재배치용)
    """
    book = None

    try:
        with ExecutionTimer() as timer:
            # 워크북 연결
            book = get_or_open_workbook(
                file_path=file_path, workbook_name=workbook_name, use_active=use_active, visible=visible
            )

            # 시트 가져오기
            target_sheet = get_sheet(book, sheet)

            # 도형 정보 수집
            shapes_info = get_shapes_info(target_sheet)

            # 타입 필터링
            if filter_type:
                filtered_shapes = []
                for shape_info in shapes_info:
                    # 타입 이름으로 필터링 (대소문자 무시)
                    if filter_type.lower() in str(shape_info.get("type", "")).lower():
                        filtered_shapes.append(shape_info)
                shapes_info = filtered_shapes

            # 상세 정보 추가 수집 (Windows에서만)
            if (detailed or include_text) and platform.system() == "Windows":
                try:
                    for i, shape_info in enumerate(shapes_info):
                        shape_name = shape_info["name"]

                        # 실제 도형 객체 찾기
                        shape_obj = None
                        for shape in target_sheet.shapes:
                            if shape.name == shape_name:
                                shape_obj = shape
                                break

                        if shape_obj:
                            # 텍스트 정보 추가
                            if include_text:
                                try:
                                    if hasattr(shape_obj.api, "TextFrame"):
                                        if shape_obj.api.TextFrame.HasText:
                                            shape_info["text_content"] = shape_obj.api.TextFrame.Characters().Text
                                        else:
                                            shape_info["text_content"] = ""
                                except Exception:
                                    shape_info["text_content"] = None

                            # 추가 상세 정보
                            if detailed:
                                try:
                                    # 그림자 정보
                                    if hasattr(shape_obj.api, "Shadow"):
                                        shape_info["has_shadow"] = shape_obj.api.Shadow.Type != 0
                                        if shape_info["has_shadow"]:
                                            from .utils import rgb_to_hex

                                            shape_info["shadow_color"] = rgb_to_hex(shape_obj.api.Shadow.ForeColor.RGB)

                                    # 테두리 정보
                                    if hasattr(shape_obj.api, "Line"):
                                        shape_info["has_border"] = shape_obj.api.Line.Visible
                                        if shape_info["has_border"]:
                                            shape_info["border_color"] = rgb_to_hex(shape_obj.api.Line.ForeColor.RGB)

                                except Exception:
                                    pass

                except Exception:
                    # 상세 정보 수집 실패해도 기본 정보는 반환
                    pass

            # 응답 데이터 구성
            response_data = {
                "shapes": shapes_info,
                "total_shapes": len(shapes_info),
                "sheet": target_sheet.name,
                "workbook": normalize_path(book.name),
                "options": {"detailed": detailed, "include_text": include_text, "filter_type": filter_type},
            }

            # 플랫폼별 기능 지원 정보
            if platform.system() != "Windows":
                response_data["note"] = "macOS에서는 제한된 도형 정보만 제공됩니다"

            # 통계 정보
            if shapes_info:
                response_data["statistics"] = {
                    "unique_types": len(set(str(s.get("type", "unknown")) for s in shapes_info)),
                    "has_text_count": sum(1 for s in shapes_info if s.get("has_text", False)),
                    "visible_count": sum(1 for s in shapes_info if s.get("visible", True)),
                }

            response = create_success_response(
                data=response_data,
                command="shape-list",
                message=f"{len(shapes_info)}개의 도형 정보를 조회했습니다",
                execution_time_ms=timer.execution_time_ms,
                book=book,
                shapes_count=len(shapes_info),
            )

            print(json.dumps(response, ensure_ascii=False, indent=2))

    except Exception as e:
        error_response = create_error_response(e, "shape-list")
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


if __name__ == "__main__":
    shape_list()
