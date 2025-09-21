"""
도형 삭제 명령어
xlwings를 활용한 Excel 도형 삭제 기능
"""

import json
from typing import Optional

import typer
import xlwings as xw

from pyhub_office_automation.version import get_version

from .utils import (
    ExecutionTimer,
    create_error_response,
    create_success_response,
    get_or_open_workbook,
    get_shape_by_name,
    get_sheet,
    normalize_path,
)


def shape_delete(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="도형을 삭제할 Excel 파일의 절대 경로"),
    use_active: bool = typer.Option(False, "--use-active", help="현재 활성 워크북 사용"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help='열린 워크북 이름으로 접근 (예: "Sales.xlsx")'),
    sheet: Optional[str] = typer.Option(None, "--sheet", help="도형을 삭제할 시트 이름 (지정하지 않으면 활성 시트)"),
    shape_name: Optional[str] = typer.Option(None, "--shape-name", help="삭제할 도형 이름"),
    shape_index: Optional[int] = typer.Option(None, "--shape-index", help="삭제할 도형 인덱스 (1부터 시작)"),
    all_shapes: bool = typer.Option(False, "--all-shapes", help="시트의 모든 도형 삭제 (위험: 확인 없이 삭제됨)"),
    confirm_all: bool = typer.Option(False, "--confirm-all", help="모든 도형 삭제 시 확인 (--all-shapes와 함께 사용)"),
    dry_run: bool = typer.Option(False, "--dry-run", help="실제 삭제하지 않고 삭제 대상만 확인"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택 (json/text)"),
    visible: bool = typer.Option(False, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부 (기본값: False)"),
    save: bool = typer.Option(True, "--save", help="삭제 후 파일 저장 여부 (기본값: True)"),
):
    """
    Excel 시트에서 도형을 삭제합니다.

    이름 또는 인덱스로 특정 도형을 삭제하거나, 시트의 모든 도형을 일괄 삭제할 수 있습니다.
    대시보드 정리 및 레이아웃 재구성 시 유용합니다.

    === 워크북 접근 방법 ===
    - --file-path: 파일 경로로 워크북 열기
    - --use-active: 현재 활성 워크북 사용
    - --workbook-name: 열린 워크북 이름으로 접근 (예: "Sales.xlsx")

    === 삭제 방법 ===
    1. 이름으로 삭제: --shape-name "TitleArea"
    2. 인덱스로 삭제: --shape-index 1 (첫 번째 도형)
    3. 전체 삭제: --all-shapes (위험!)

    === 안전 기능 ===
    • --dry-run: 실제 삭제하지 않고 대상만 확인
    • --confirm-all: 전체 삭제 시 추가 확인 (권장)
    • 자동 백업: 파일 저장으로 변경사항 보존

    === 사용 시나리오 ===

    # 1. 특정 도형 삭제 (이름으로)
    oa excel shape-delete --use-active --shape-name "TitleArea"

    # 2. 첫 번째 도형 삭제 (인덱스로)
    oa excel shape-delete --use-active --shape-index 1

    # 3. 삭제 전 확인 (dry-run)
    oa excel shape-delete --use-active --shape-name "OldChart" --dry-run

    # 4. 시트의 모든 도형 삭제 (주의!)
    oa excel shape-delete --use-active --all-shapes --confirm-all

    # 5. 특정 시트의 도형 삭제
    oa excel shape-delete --file-path "dashboard.xlsx" --sheet "OldDashboard" \\
        --shape-name "ObsoleteShape" --save true

    === 대시보드 관리 활용 ===
    • 구버전 도형 정리
    • 테스트 도형 제거
    • 레이아웃 재구성 전 정리
    • 중복 도형 제거
    • 대시보드 템플릿 정리

    === 주의사항 ===
    • 삭제된 도형은 복구할 수 없습니다
    • --all-shapes는 매우 위험하므로 신중히 사용하세요
    • 중요한 파일은 사전에 백업하세요
    • --dry-run으로 먼저 확인하는 것을 권장합니다
    """
    book = None

    try:
        with ExecutionTimer() as timer:
            # 삭제 방법 검증
            delete_methods = sum([bool(shape_name), bool(shape_index is not None), bool(all_shapes)])

            if delete_methods == 0:
                raise ValueError("삭제 방법을 지정해야 합니다: --shape-name, --shape-index, 또는 --all-shapes")
            elif delete_methods > 1:
                raise ValueError("삭제 방법은 하나만 지정할 수 있습니다")

            # 전체 삭제 안전 확인
            if all_shapes and not confirm_all and not dry_run:
                raise ValueError("모든 도형을 삭제하려면 --confirm-all 플래그를 함께 사용하세요")

            # 워크북 연결
            book = get_or_open_workbook(
                file_path=file_path, workbook_name=workbook_name, use_active=use_active, visible=visible
            )

            # 시트 가져오기
            target_sheet = get_sheet(book, sheet)

            # 현재 도형 개수 확인
            initial_shape_count = len(target_sheet.shapes)
            if initial_shape_count == 0:
                raise ValueError("시트에 삭제할 도형이 없습니다")

            # 삭제 대상 확인
            shapes_to_delete = []
            deleted_info = []

            if shape_name:
                # 이름으로 삭제
                shape_obj = get_shape_by_name(target_sheet, shape_name)
                if not shape_obj:
                    raise ValueError(f"도형 '{shape_name}'을 찾을 수 없습니다")

                shapes_to_delete.append(shape_obj)
                deleted_info.append(
                    {
                        "name": shape_obj.name,
                        "method": "name",
                        "position": {"left": getattr(shape_obj, "left", 0), "top": getattr(shape_obj, "top", 0)},
                        "size": {"width": getattr(shape_obj, "width", 0), "height": getattr(shape_obj, "height", 0)},
                    }
                )

            elif shape_index is not None:
                # 인덱스로 삭제
                if shape_index < 1 or shape_index > initial_shape_count:
                    raise ValueError(f"도형 인덱스가 범위를 벗어났습니다 (1-{initial_shape_count})")

                shape_obj = target_sheet.shapes[shape_index - 1]  # 0-based index
                shapes_to_delete.append(shape_obj)
                deleted_info.append(
                    {
                        "name": shape_obj.name,
                        "method": "index",
                        "index": shape_index,
                        "position": {"left": getattr(shape_obj, "left", 0), "top": getattr(shape_obj, "top", 0)},
                        "size": {"width": getattr(shape_obj, "width", 0), "height": getattr(shape_obj, "height", 0)},
                    }
                )

            elif all_shapes:
                # 모든 도형 삭제
                for i, shape_obj in enumerate(target_sheet.shapes):
                    shapes_to_delete.append(shape_obj)
                    deleted_info.append(
                        {
                            "name": shape_obj.name,
                            "method": "all",
                            "index": i + 1,
                            "position": {"left": getattr(shape_obj, "left", 0), "top": getattr(shape_obj, "top", 0)},
                            "size": {"width": getattr(shape_obj, "width", 0), "height": getattr(shape_obj, "height", 0)},
                        }
                    )

            # Dry-run 모드
            if dry_run:
                response_data = {
                    "dry_run": True,
                    "shapes_to_delete": deleted_info,
                    "total_to_delete": len(shapes_to_delete),
                    "current_shape_count": initial_shape_count,
                    "remaining_after_delete": initial_shape_count - len(shapes_to_delete),
                    "sheet": target_sheet.name,
                    "workbook": normalize_path(book.name),
                }

                response = create_success_response(
                    data=response_data,
                    command="shape-delete",
                    message=f"[DRY RUN] {len(shapes_to_delete)}개의 도형이 삭제될 예정입니다",
                    execution_time_ms=timer.execution_time_ms,
                    book=book,
                )

                typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
                raise typer.Exit(0)

            # 실제 삭제 수행
            deleted_count = 0
            for shape_obj in shapes_to_delete:
                try:
                    shape_obj.delete()
                    deleted_count += 1
                except Exception as e:
                    # 개별 도형 삭제 실패는 경고로 처리
                    for info in deleted_info:
                        if info.get("name") == shape_obj.name:
                            info["delete_error"] = str(e)
                            break

            # 파일 저장
            if save and file_path:
                book.save()

            # 결과 확인
            final_shape_count = len(target_sheet.shapes)

            # 성공 응답 생성
            response_data = {
                "deleted_shapes": deleted_info,
                "total_deleted": deleted_count,
                "initial_shape_count": initial_shape_count,
                "final_shape_count": final_shape_count,
                "sheet": target_sheet.name,
                "workbook": normalize_path(book.name),
            }

            # 삭제 실패한 도형이 있는 경우 정보 추가
            failed_deletes = [info for info in deleted_info if "delete_error" in info]
            if failed_deletes:
                response_data["failed_deletes"] = failed_deletes
                response_data["failed_count"] = len(failed_deletes)

            message = f"{deleted_count}개의 도형이 성공적으로 삭제되었습니다"
            if failed_deletes:
                message += f" ({len(failed_deletes)}개 실패)"

            response = create_success_response(
                data=response_data,
                command="shape-delete",
                message=message,
                execution_time_ms=timer.execution_time_ms,
                book=book,
                shapes_deleted=deleted_count,
            )

            typer.echo(json.dumps(response, ensure_ascii=False, indent=2))

    except Exception as e:
        error_response = create_error_response(e, "shape-delete")
        typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        raise typer.Exit(1)

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

    raise typer.Exit(0)
