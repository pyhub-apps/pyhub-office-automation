"""
슬라이서 추가 명령어
xlwings를 활용한 Excel 슬라이서 생성 기능
대시보드 필터링 및 상호작용 구성
"""

import json
import platform
from typing import Optional

import typer
import xlwings as xw

from pyhub_office_automation.version import get_version

from .engines import get_engine
from .utils import (
    ExecutionTimer,
    OutputFormat,
    SlicerStyle,
    check_slicer_cache_conflicts,
    create_error_response,
    create_success_response,
    generate_unique_slicer_name,
    get_or_open_workbook,
    get_pivot_tables,
    get_sheet,
    get_slicer_cache_by_field,
    normalize_path,
    remove_slicer_cache,
    validate_slicer_position,
)


def slicer_add(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="슬라이서를 추가할 Excel 파일의 절대 경로"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help='열린 워크북 이름으로 접근 (예: "Sales.xlsx")'),
    sheet: Optional[str] = typer.Option(None, "--sheet", help="슬라이서를 배치할 시트 이름 (지정하지 않으면 활성 시트)"),
    pivot_table: str = typer.Option(..., "--pivot-table", help="슬라이서를 생성할 피벗테이블 이름"),
    field: str = typer.Option(..., "--field", help="슬라이서로 만들 피벗테이블 필드 이름"),
    left: int = typer.Option(100, "--left", help="슬라이서의 왼쪽 위치 (픽셀, 기본값: 100)"),
    top: int = typer.Option(100, "--top", help="슬라이서의 위쪽 위치 (픽셀, 기본값: 100)"),
    width: int = typer.Option(200, "--width", help="슬라이서의 너비 (픽셀, 기본값: 200)"),
    height: int = typer.Option(150, "--height", help="슬라이서의 높이 (픽셀, 기본값: 150)"),
    name: Optional[str] = typer.Option(None, "--name", help="슬라이서 이름 (지정하지 않으면 자동 생성)"),
    caption: Optional[str] = typer.Option(None, "--caption", help="슬라이서 제목 (지정하지 않으면 필드명 사용)"),
    style: SlicerStyle = typer.Option(SlicerStyle.LIGHT, "--style", help="슬라이서 스타일 (light/medium/dark, 기본값: light)"),
    columns: int = typer.Option(1, "--columns", help="슬라이서 항목 열 개수 (기본값: 1)"),
    item_height: Optional[int] = typer.Option(None, "--item-height", help="슬라이서 항목 높이 (픽셀)"),
    show_header: bool = typer.Option(True, "--show-header", help="슬라이서 헤더 표시 (기본값: True)"),
    force: bool = typer.Option(False, "--force", help="기존 SlicerCache 제거 후 재생성 (기본값: False)"),
    reuse_cache: bool = typer.Option(
        False, "--reuse-cache", help="기존 SlicerCache 재사용하여 새 슬라이서 추가 (기본값: False)"
    ),
    output_format: OutputFormat = typer.Option(OutputFormat.JSON, "--format", help="출력 형식 선택 (json/text)"),
    visible: bool = typer.Option(False, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부 (기본값: False)"),
    save: bool = typer.Option(True, "--save", help="생성 후 파일 저장 여부 (기본값: True)"),
):
    """
    📊 Excel 피벗테이블 기반 슬라이서를 생성합니다.

    피벗테이블의 특정 필드를 슬라이서로 만들어 대화형 대시보드를 구성할 수 있습니다.
    여러 피벗테이블에 연결하여 통합 필터링 기능을 제공합니다.

    ## 📁 워크북 접근 방법

    - `--file-path`: 파일 경로로 워크북 열기
    - `--workbook-name`: 열린 워크북 이름으로 접근 (예: "Sales.xlsx")

    ## ✅ 슬라이서 생성 조건

    - 대상 피벗테이블이 존재해야 함
    - 지정한 필드가 피벗테이블에 포함되어 있어야 함
    - Windows에서만 완전 지원 (macOS 제한)

    ## 🔧 중복 해결 옵션 (Issue #71)

    - `--force`: 기존 SlicerCache 제거 후 재생성
    - `--reuse-cache`: 기존 SlicerCache에 새 슬라이서 추가

    **💡 Tip**: 동일한 필드에 대한 SlicerCache가 이미 존재할 때 사용

    ## 🚀 사용 예시

    **기본 사용법:**
    ```bash
    oa excel slicer-add --pivot-table "SalesPivot" --field "지역"
    ```

    **중복 해결:**
    ```bash
    # 강제 재생성
    oa excel slicer-add --pivot-table "SalesPivot" --field "지역" --force

    # 기존 캐시 재사용
    oa excel slicer-add --pivot-table "SalesPivot" --field "지역" --reuse-cache
    ```

    ## ⚠️ 주의사항

    - Windows에서만 모든 기능 지원
    - 피벗테이블이 존재하지 않으면 생성 불가
    - 필드명은 피벗테이블에 실제 존재하는 이름 사용
    """
    book = None

    try:
        # Enum 타입이므로 별도 검증 불필요

        with ExecutionTimer() as timer:
            # Windows 플랫폼 확인
            if platform.system() != "Windows":
                raise RuntimeError("슬라이서는 Windows에서만 지원됩니다")

            # 슬라이서 위치와 크기 검증
            is_valid, error_msg = validate_slicer_position(left, top, width, height)
            if not is_valid:
                raise ValueError(error_msg)

            # 워크북 연결
            book = get_or_open_workbook(file_path=file_path, workbook_name=workbook_name, visible=visible)

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
                    f"필드 '{field}'를 피벗테이블에서 찾을 수 없습니다. " f"사용 가능한 필드: {', '.join(available_fields)}"
                )

            # SlicerCache 충돌 확인 및 처리 (Issue #71)
            conflict_info = check_slicer_cache_conflicts(book, pivot_table, field)
            existing_slicer_cache = None

            if conflict_info["has_conflict"]:
                if force:
                    # 기존 캐시 제거 후 재생성
                    if remove_slicer_cache(book, conflict_info["existing_cache"]):
                        # 제거 성공, 계속 진행
                        pass
                    else:
                        raise RuntimeError(f"기존 SlicerCache 제거에 실패했습니다")
                elif reuse_cache:
                    # 기존 캐시 재사용
                    existing_slicer_cache = conflict_info["existing_cache"]
                else:
                    # 충돌 시 명확한 안내 메시지
                    options_msg = "\n".join([f"  • {opt}" for opt in conflict_info["resolution_options"]])
                    raise ValueError(
                        f"{conflict_info['message']}\n\n"
                        f"해결 방법:\n{options_msg}\n"
                        f"  • 기존 슬라이서 확인: oa excel slicer-list"
                    )

            # 옵션 충돌 검사
            if force and reuse_cache:
                raise ValueError("--force와 --reuse-cache 옵션은 동시에 사용할 수 없습니다")

            # 슬라이서 이름 결정
            if not name:
                name = generate_unique_slicer_name(book, f"{field}Slicer")

            # 캡션 결정
            if not caption:
                caption = field

            # 슬라이서 생성 (Engine Layer 사용)
            try:
                # Engine 가져오기
                engine = get_engine()

                # 슬라이서 추가 옵션 준비
                kwargs = {
                    "columns": columns,
                    "style": style.value if hasattr(style, "value") else style,
                    "caption": caption,
                    "show_header": show_header,
                }

                if item_height:
                    kwargs["item_height"] = item_height

                # 기존 캐시 재사용 처리
                if existing_slicer_cache:
                    kwargs["existing_cache"] = existing_slicer_cache

                # Engine 메서드로 슬라이서 추가
                result = engine.add_slicer(
                    workbook=book.api,
                    sheet_name=target_sheet.name,
                    pivot_table_name=pivot_table,
                    field_name=field,
                    left=left,
                    top=top,
                    width=width,
                    height=height,
                    slicer_name=name,
                    **kwargs,
                )

                # result에서 슬라이서 정보 추출
                slicer_name = result.get("name", name)
                slicer_items = result.get("slicer_items", [])

            except Exception as e:
                raise RuntimeError(f"슬라이서 생성 실패: {str(e)}")

            # 파일 저장
            if save and file_path:
                book.save()

            # 성공 응답 생성
            response_data = {
                "slicer_name": slicer_name,
                "slicer_caption": caption,
                "pivot_table": pivot_table,
                "field": field,
                "position": {"left": left, "top": top},
                "size": {"width": width, "height": height},
                "settings": {
                    "style": style.value if hasattr(style, "value") else style,
                    "columns": columns,
                    "show_header": show_header,
                },
                "slicer_items": slicer_items,
                "total_items": len(slicer_items),
                "sheet": target_sheet.name,
                "workbook": normalize_path(book.name),
                "cache_action": "reused" if existing_slicer_cache else "created",
                "conflict_resolved": conflict_info["has_conflict"],
            }

            if item_height:
                response_data["settings"]["item_height"] = item_height

            # 메시지 생성
            if existing_slicer_cache:
                message = f"기존 SlicerCache를 재사용하여 슬라이서 '{name}'을 추가했습니다 ({len(slicer_items)}개 항목)"
            elif conflict_info["has_conflict"] and force:
                message = f"기존 SlicerCache를 제거하고 슬라이서 '{name}'을 재생성했습니다 ({len(slicer_items)}개 항목)"
            else:
                message = f"슬라이서 '{name}'이 성공적으로 생성되었습니다 ({len(slicer_items)}개 항목)"

            response = create_success_response(
                data=response_data,
                command="slicer-add",
                message=message,
                execution_time_ms=timer.execution_time_ms,
                book=book,
                slicer_items=len(slicer_items),
            )

            print(json.dumps(response, ensure_ascii=False, indent=2))

    except Exception as e:
        error_response = create_error_response(e, "slicer-add")
        print(json.dumps(error_response, ensure_ascii=False, indent=2))
        return 1

    finally:
        # 새로 생성한 워크북인 경우에만 정리
        if book and file_path and not workbook_name:
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
    slicer_add()
