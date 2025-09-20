"""
슬라이서 추가 명령어
xlwings를 활용한 Excel 슬라이서 생성 기능
대시보드 필터링 및 상호작용 구성
"""

import json
import platform
import click
import xlwings as xw
from ..version import get_version
from .utils import (
    get_or_open_workbook, get_sheet, create_error_response, create_success_response,
    ExecutionTimer, get_pivot_tables, validate_slicer_position,
    generate_unique_slicer_name, normalize_path
)


@click.command()
@click.option('--file-path',
              help='슬라이서를 추가할 Excel 파일의 절대 경로')
@click.option('--use-active', is_flag=True,
              help='현재 활성 워크북 사용')
@click.option('--workbook-name',
              help='열린 워크북 이름으로 접근 (예: "Sales.xlsx")')
@click.option('--sheet',
              help='슬라이서를 배치할 시트 이름 (지정하지 않으면 활성 시트)')
@click.option('--pivot-table', required=True,
              help='슬라이서를 생성할 피벗테이블 이름')
@click.option('--field', required=True,
              help='슬라이서로 만들 피벗테이블 필드 이름')
@click.option('--left', type=int, default=100,
              help='슬라이서의 왼쪽 위치 (픽셀, 기본값: 100)')
@click.option('--top', type=int, default=100,
              help='슬라이서의 위쪽 위치 (픽셀, 기본값: 100)')
@click.option('--width', type=int, default=200,
              help='슬라이서의 너비 (픽셀, 기본값: 200)')
@click.option('--height', type=int, default=150,
              help='슬라이서의 높이 (픽셀, 기본값: 150)')
@click.option('--name',
              help='슬라이서 이름 (지정하지 않으면 자동 생성)')
@click.option('--caption',
              help='슬라이서 제목 (지정하지 않으면 필드명 사용)')
@click.option('--style',
              type=click.Choice(['light', 'medium', 'dark']),
              default='light',
              help='슬라이서 스타일 (기본값: light)')
@click.option('--columns', type=int, default=1,
              help='슬라이서 항목 열 개수 (기본값: 1)')
@click.option('--item-height', type=int,
              help='슬라이서 항목 높이 (픽셀)')
@click.option('--show-header', is_flag=True, default=True,
              help='슬라이서 헤더 표시 (기본값: True)')
@click.option('--format', 'output_format', default='json',
              type=click.Choice(['json', 'text']),
              help='출력 형식 선택')
@click.option('--visible', default=False, type=bool,
              help='Excel 애플리케이션을 화면에 표시할지 여부 (기본값: False)')
@click.option('--save', default=True, type=bool,
              help='생성 후 파일 저장 여부 (기본값: True)')
@click.version_option(version=get_version(), prog_name="oa excel slicer-add")
def slicer_add(file_path, use_active, workbook_name, sheet, pivot_table, field,
               left, top, width, height, name, caption, style, columns, item_height,
               show_header, output_format, visible, save):
    """
    Excel 피벗테이블 기반 슬라이서를 생성합니다.

    피벗테이블의 특정 필드를 슬라이서로 만들어 대화형 대시보드 구성이 가능하며,
    여러 피벗테이블에 연결하여 통합 필터링 기능을 제공합니다.

    === 워크북 접근 방법 ===
    - --file-path: 파일 경로로 워크북 열기
    - --use-active: 현재 활성 워크북 사용
    - --workbook-name: 열린 워크북 이름으로 접근 (예: "Sales.xlsx")

    === 슬라이서 생성 조건 ===
    • 대상 피벗테이블이 존재해야 함
    • 지정한 필드가 피벗테이블에 포함되어 있어야 함
    • Windows에서만 완전 지원 (macOS 제한)

    === 슬라이서 설정 옵션 ===
    • --pivot-table: 대상 피벗테이블 이름
    • --field: 슬라이서로 만들 필드명
    • --left, --top: 슬라이서 위치 (픽셀)
    • --width, --height: 슬라이서 크기 (픽셀)
    • --name: 슬라이서 고유 이름
    • --caption: 사용자에게 표시될 제목

    === 스타일 및 레이아웃 ===
    • --style: light, medium, dark 스타일
    • --columns: 항목을 표시할 열 개수
    • --item-height: 각 항목의 높이
    • --show-header: 헤더(제목) 표시 여부

    === 대시보드 구성 시나리오 ===

    # 1. 지역별 매출 필터 슬라이서
    oa excel slicer-add --use-active --pivot-table "SalesPivot" --field "지역" \\
        --left 90 --top 400 --width 200 --height 120 \\
        --name "RegionSlicer" --caption "지역 선택" --columns 2

    # 2. 기간 필터 슬라이서 (세로 레이아웃)
    oa excel slicer-add --use-active --pivot-table "SalesPivot" --field "월" \\
        --left 320 --top 400 --width 150 --height 180 \\
        --name "MonthSlicer" --caption "기간" --columns 1 --style "medium"

    # 3. 제품 카테고리 슬라이서 (가로 레이아웃)
    oa excel slicer-add --use-active --pivot-table "ProductPivot" --field "카테고리" \\
        --left 500 --top 400 --width 300 --height 80 \\
        --name "CategorySlicer" --caption "제품 분류" --columns 3 --item-height 25

    # 4. 판매자 필터 (어두운 스타일)
    oa excel slicer-add --use-active --pivot-table "SalesPivot" --field "판매자" \\
        --left 90 --top 550 --width 180 --height 150 \\
        --name "SalespersonSlicer" --caption "담당자" --style "dark"

    === 고급 대시보드 구성 ===

    # 다중 피벗테이블 연동 준비 (연결은 slicer-connect로)
    # 1. 메인 매출 분석용
    oa excel slicer-add --use-active --pivot-table "MainSalesPivot" --field "지역" \\
        --left 100 --top 500 --width 200 --height 100 --name "MainRegionSlicer"

    # 2. 트렌드 분석용 (같은 필드, 다른 피벗테이블)
    oa excel slicer-add --use-active --pivot-table "TrendPivot" --field "지역" \\
        --left 320 --top 500 --width 200 --height 100 --name "TrendRegionSlicer"

    # 3. 통합 슬라이서로 업그레이드 예정
    # 이후 slicer-connect로 두 피벗테이블을 하나의 슬라이서에 연결

    === 슬라이서 배치 가이드 ===

    # 뉴모피즘 슬라이서 박스 내부 배치
    # 1. 배경 도형 먼저 생성 (shape-add로)
    oa excel shape-add --use-active --shape-type rounded_rectangle \\
        --left 80 --top 380 --width 740 --height 140 \\
        --style-preset slicer-box --name "SlicerBackground"

    # 2. 슬라이서들을 배경 내부에 배치
    oa excel slicer-add --use-active --pivot-table "SalesPivot" --field "지역" \\
        --left 100 --top 400 --width 150 --height 100
    oa excel slicer-add --use-active --pivot-table "SalesPivot" --field "월" \\
        --left 270 --top 400 --width 150 --height 100
    oa excel slicer-add --use-active --pivot-table "SalesPivot" --field "제품" \\
        --left 440 --top 400 --width 150 --height 100
    oa excel slicer-add --use-active --pivot-table "SalesPivot" --field "담당자" \\
        --left 610 --top 400 --width 150 --height 100

    === 사용 팁 ===
    • 슬라이서 이름은 추후 연결 및 관리를 위해 명확하게 지정
    • caption은 사용자 친화적인 한글 제목 권장
    • 항목이 많은 필드는 columns를 늘려 공간 효율성 확보
    • 대시보드 스타일에 맞는 style 선택
    • 슬라이서 간 일정한 간격 유지로 정돈된 레이아웃 구성

    === 주의사항 ===
    • Windows에서만 모든 기능 지원
    • 피벗테이블이 존재하지 않으면 생성 불가
    • 필드명은 피벗테이블에 실제 존재하는 이름 사용
    • 슬라이서 이름 중복 시 자동으로 숫자 추가
    """
    book = None

    try:
        with ExecutionTimer() as timer:
            # Windows 플랫폼 확인
            if platform.system() != "Windows":
                raise RuntimeError("슬라이서는 Windows에서만 지원됩니다")

            # 슬라이서 위치와 크기 검증
            is_valid, error_msg = validate_slicer_position(left, top, width, height)
            if not is_valid:
                raise ValueError(error_msg)

            # 워크북 연결
            book = get_or_open_workbook(
                file_path=file_path,
                workbook_name=workbook_name,
                use_active=use_active,
                visible=visible
            )

            # 시트 가져오기
            target_sheet = get_sheet(book, sheet)

            # 피벗테이블 존재 확인
            pivot_tables = get_pivot_tables(target_sheet)
            target_pivot = None

            for pt in pivot_tables:
                if pt["name"] == pivot_table:
                    target_pivot = pt
                    break

            if not target_pivot:
                available_pivots = [pt["name"] for pt in pivot_tables]
                if available_pivots:
                    raise ValueError(
                        f"피벗테이블 '{pivot_table}'을 찾을 수 없습니다. "
                        f"사용 가능한 피벗테이블: {', '.join(available_pivots)}"
                    )
                else:
                    raise ValueError("시트에 피벗테이블이 없습니다")

            # 필드 존재 확인
            available_fields = [f["name"] for f in target_pivot["fields"]]
            if field not in available_fields:
                raise ValueError(
                    f"필드 '{field}'를 피벗테이블에서 찾을 수 없습니다. "
                    f"사용 가능한 필드: {', '.join(available_fields)}"
                )

            # 슬라이서 이름 결정
            if not name:
                name = generate_unique_slicer_name(book, f"{field}Slicer")

            # 캡션 결정
            if not caption:
                caption = field

            # 슬라이서 생성 (Windows COM API)
            try:
                # 피벗테이블 객체 가져오기
                pivot_table_obj = None
                for pt in target_sheet.api.PivotTables():
                    if pt.Name == pivot_table:
                        pivot_table_obj = pt
                        break

                if not pivot_table_obj:
                    raise RuntimeError(f"피벗테이블 '{pivot_table}' 객체를 가져올 수 없습니다")

                # 필드 객체 가져오기
                field_obj = None
                try:
                    field_obj = pivot_table_obj.PivotFields(field)
                except:
                    raise ValueError(f"필드 '{field}'에 접근할 수 없습니다")

                # 슬라이서 생성
                slicer_cache = book.api.SlicerCaches.Add(
                    Source=pivot_table_obj,
                    SourceField=field_obj
                )

                # 슬라이서 이름 설정
                slicer_cache.Name = name

                # 슬라이서 배치
                slicer = slicer_cache.Slicers.Add(
                    SlicerDestination=target_sheet.api,
                    Left=left,
                    Top=top,
                    Width=width,
                    Height=height
                )

                # 슬라이서 설정
                slicer.Caption = caption

                # 스타일 설정
                style_map = {
                    'light': 'SlicerStyleLight1',
                    'medium': 'SlicerStyleLight2',
                    'dark': 'SlicerStyleDark1'
                }

                try:
                    if style in style_map:
                        slicer.Style = style_map[style]
                except:
                    pass

                # 레이아웃 설정
                if columns > 1:
                    try:
                        slicer.NumberOfColumns = columns
                    except:
                        pass

                if item_height:
                    try:
                        slicer.RowHeight = item_height
                    except:
                        pass

                # 헤더 표시 설정
                try:
                    slicer.DisplayHeader = show_header
                except:
                    pass

            except Exception as e:
                raise RuntimeError(f"슬라이서 생성 실패: {str(e)}")

            # 파일 저장
            if save and file_path:
                book.save()

            # 슬라이서 정보 수집
            slicer_items = []
            try:
                for item in slicer_cache.SlicerItems():
                    slicer_items.append({
                        "name": item.Name,
                        "selected": item.Selected
                    })
            except:
                pass

            # 성공 응답 생성
            response_data = {
                "slicer_name": name,
                "slicer_caption": caption,
                "pivot_table": pivot_table,
                "field": field,
                "position": {
                    "left": left,
                    "top": top
                },
                "size": {
                    "width": width,
                    "height": height
                },
                "settings": {
                    "style": style,
                    "columns": columns,
                    "show_header": show_header
                },
                "slicer_items": slicer_items,
                "total_items": len(slicer_items),
                "sheet": target_sheet.name,
                "workbook": normalize_path(book.name)
            }

            if item_height:
                response_data["settings"]["item_height"] = item_height

            message = f"슬라이서 '{name}'이 성공적으로 생성되었습니다 ({len(slicer_items)}개 항목)"

            response = create_success_response(
                data=response_data,
                command="slicer-add",
                message=message,
                execution_time_ms=timer.execution_time_ms,
                book=book,
                slicer_items=len(slicer_items)
            )

            print(json.dumps(response, ensure_ascii=False, indent=2))

    except Exception as e:
        error_response = create_error_response(e, "slicer-add")
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


if __name__ == '__main__':
    slicer_add()