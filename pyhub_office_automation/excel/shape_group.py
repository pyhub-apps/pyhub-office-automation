"""
도형 그룹화 명령어
xlwings를 활용한 Excel 도형 그룹화 및 해제 기능
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
    generate_unique_shape_name,
    get_or_open_workbook,
    get_shape_by_name,
    get_sheet,
    normalize_path,
)


def shape_group(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="도형을 그룹화할 Excel 파일의 절대 경로"),
    use_active: bool = typer.Option(False, "--use-active", help="현재 활성 워크북 사용"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help='열린 워크북 이름으로 접근 (예: "Sales.xlsx")'),
    sheet: Optional[str] = typer.Option(None, "--sheet", help="도형이 있는 시트 이름 (지정하지 않으면 활성 시트)"),
    action: str = typer.Option(..., "--action", help="작업 유형: group(그룹화) 또는 ungroup(그룹 해제)"),
    shape_names: Optional[str] = typer.Option(
        None, "--shape-names", help='그룹화할 도형 이름들 (쉼표로 구분, 예: "Shape1,Shape2,Shape3")'
    ),
    group_name: Optional[str] = typer.Option(None, "--group-name", help="그룹 이름 (그룹화 시, 지정하지 않으면 자동 생성)"),
    target_group: Optional[str] = typer.Option(None, "--target-group", help="해제할 그룹 이름 (ungroup 시 필수)"),
    all_groups: bool = typer.Option(False, "--all-groups", help="시트의 모든 그룹 해제 (ungroup 시)"),
    include_nested: bool = typer.Option(False, "--include-nested", help="중첩된 그룹도 포함하여 처리"),
    dry_run: bool = typer.Option(False, "--dry-run", help="실제 작업하지 않고 대상만 확인"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택 (json/text)"),
    visible: bool = typer.Option(False, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부 (기본값: False)"),
    save: bool = typer.Option(True, "--save", help="작업 후 파일 저장 여부 (기본값: True)"),
):
    """
    Excel 도형을 그룹화하거나 그룹을 해제합니다.

    여러 도형을 하나의 그룹으로 묶어 일괄 이동, 크기 조정, 스타일링이 가능하며,
    대시보드 레이아웃 관리와 복잡한 도형 구조 정리에 유용합니다.

    === 워크북 접근 방법 ===
    - --file-path: 파일 경로로 워크북 열기
    - --use-active: 현재 활성 워크북 사용
    - --workbook-name: 열린 워크북 이름으로 접근 (예: "Sales.xlsx")

    === 그룹화 작업 (--action group) ===
    • --shape-names: 그룹화할 도형들 (쉼표 구분)
    • --group-name: 그룹 이름 (선택, 자동 생성 가능)
    • 최소 2개 이상의 도형 필요

    === 그룹 해제 작업 (--action ungroup) ===
    • --target-group: 해제할 특정 그룹 이름
    • --all-groups: 시트의 모든 그룹 해제
    • --include-nested: 중첩 그룹까지 재귀적 해제

    === 안전 기능 ===
    • --dry-run: 실제 작업 전 대상 확인
    • 자동 이름 생성으로 중복 방지
    • 작업 전 검증으로 오류 최소화

    === 그룹화 시나리오 ===

    # 1. 차트 영역 구성 요소들 그룹화
    oa excel shape-group --use-active --action group \\
        --shape-names "ChartArea,ChartTitle,ChartLegend" \\
        --group-name "Chart1Group"

    # 2. 대시보드 헤더 영역 그룹화
    oa excel shape-group --use-active --action group \\
        --shape-names "TitleBackground,MainTitle,SubTitle,DateLabel" \\
        --group-name "HeaderGroup"

    # 3. KPI 카드들 그룹화
    oa excel shape-group --use-active --action group \\
        --shape-names "KPI1Card,KPI1Label,KPI1Value,KPI2Card,KPI2Label,KPI2Value" \\
        --group-name "KPIGroup"

    # 4. 슬라이서 영역 그룹화
    oa excel shape-group --use-active --action group \\
        --shape-names "SlicerBackground,Slicer1,Slicer2,SlicerTitle" \\
        --group-name "FilterGroup"

    === 그룹 해제 시나리오 ===

    # 1. 특정 그룹 해제
    oa excel shape-group --use-active --action ungroup --target-group "Chart1Group"

    # 2. 모든 그룹 해제 (레이아웃 재구성 시)
    oa excel shape-group --use-active --action ungroup --all-groups

    # 3. 중첩 그룹까지 완전 해제
    oa excel shape-group --use-active --action ungroup --all-groups --include-nested

    # 4. 해제 전 확인 (dry-run)
    oa excel shape-group --use-active --action ungroup \\
        --target-group "ComplexGroup" --include-nested --dry-run

    === 대시보드 관리 활용 ===

    # 계층적 그룹 구조 생성
    # 1단계: 개별 요소들 그룹화
    oa excel shape-group --use-active --action group \\
        --shape-names "Chart1,Chart1Title" --group-name "Chart1Unit"
    oa excel shape-group --use-active --action group \\
        --shape-names "Chart2,Chart2Title" --group-name "Chart2Unit"

    # 2단계: 상위 그룹 생성
    oa excel shape-group --use-active --action group \\
        --shape-names "Chart1Unit,Chart2Unit" --group-name "ChartsSection"

    # 대시보드 전체 그룹화 (최종)
    oa excel shape-group --use-active --action group \\
        --shape-names "HeaderGroup,ChartsSection,KPIGroup,FilterGroup" \\
        --group-name "DashboardMain"

    === 그룹 관리 베스트 프랙티스 ===
    • 논리적 구조에 따른 그룹화
    • 명확한 그룹 이름 사용
    • 과도한 중첩 그룹 지양
    • 정기적인 그룹 구조 검토
    • 템플릿화를 위한 표준 그룹 명명법

    === 주의사항 ===
    • Windows에서 모든 기능 지원
    • macOS에서는 제한적 지원
    • 그룹화 후 개별 도형 접근 시 그룹 내 인덱스 필요
    • 중첩 그룹 해제 시 하위 그룹도 모두 해제됨
    """
    book = None

    try:
        # action 값 검증
        if action not in ["group", "ungroup"]:
            raise ValueError(f"action은 'group' 또는 'ungroup'이어야 합니다. 입력된 값: {action}")

        with ExecutionTimer() as timer:
            # 워크북 연결
            book = get_or_open_workbook(
                file_path=file_path, workbook_name=workbook_name, use_active=use_active, visible=visible
            )

            # 시트 가져오기
            target_sheet = get_sheet(book, sheet)

            # 작업별 검증 및 실행
            if action == "group":
                result = handle_group_action(target_sheet, shape_names, group_name, dry_run)
            else:  # ungroup
                result = handle_ungroup_action(target_sheet, target_group, all_groups, include_nested, dry_run)

            # 파일 저장
            if save and file_path and not dry_run:
                book.save()

            # 성공 응답 생성
            response_data = {
                "action": action,
                "dry_run": dry_run,
                "sheet": target_sheet.name,
                "workbook": normalize_path(book.name),
                "platform_support": "full" if platform.system() == "Windows" else "limited",
                **result,
            }

            message = result.get("message", f"{action} 작업이 완료되었습니다")
            if dry_run:
                message = f"[DRY RUN] {message}"

            response = create_success_response(
                data=response_data,
                command="shape-group",
                message=message,
                execution_time_ms=timer.execution_time_ms,
                book=book,
                action=action,
                affected_shapes=result.get("affected_shapes_count", 0),
            )

            print(json.dumps(response, ensure_ascii=False, indent=2))

    except Exception as e:
        error_response = create_error_response(e, "shape-group")
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


def handle_group_action(sheet, shape_names, group_name, dry_run):
    """그룹화 작업 처리"""
    if not shape_names:
        raise ValueError("그룹화할 도형 이름들을 지정해야 합니다 (--shape-names)")

    # 도형 이름 파싱
    shape_name_list = [name.strip() for name in shape_names.split(",") if name.strip()]

    if len(shape_name_list) < 2:
        raise ValueError("그룹화하려면 최소 2개 이상의 도형이 필요합니다")

    # 도형 객체 찾기
    shapes_to_group = []
    missing_shapes = []

    for shape_name in shape_name_list:
        shape_obj = get_shape_by_name(sheet, shape_name)
        if shape_obj:
            shapes_to_group.append(shape_obj)
        else:
            missing_shapes.append(shape_name)

    if missing_shapes:
        raise ValueError(f"다음 도형들을 찾을 수 없습니다: {', '.join(missing_shapes)}")

    # 그룹 이름 결정
    if not group_name:
        group_name = generate_unique_shape_name(sheet, "Group")

    # 기존 그룹 이름 중복 확인
    existing_group = get_shape_by_name(sheet, group_name)
    if existing_group:
        raise ValueError(f"그룹 이름 '{group_name}'이 이미 존재합니다")

    # 도형 정보 수집
    shape_details = []
    for shape_obj in shapes_to_group:
        shape_details.append(
            {
                "name": shape_obj.name,
                "position": {"left": getattr(shape_obj, "left", 0), "top": getattr(shape_obj, "top", 0)},
                "size": {"width": getattr(shape_obj, "width", 0), "height": getattr(shape_obj, "height", 0)},
            }
        )

    if dry_run:
        return {
            "shapes_to_group": shape_details,
            "group_name": group_name,
            "total_shapes": len(shapes_to_group),
            "affected_shapes_count": len(shapes_to_group),
            "message": f"{len(shapes_to_group)}개의 도형이 '{group_name}' 그룹으로 묶일 예정입니다",
        }

    # 실제 그룹화 수행 (Windows에서만 완전 지원)
    if platform.system() == "Windows":
        try:
            # xlwings에서 그룹화하려면 Shape 객체들의 Select 후 Group 호출
            # COM API를 직접 사용
            shape_range = []
            for shape_obj in shapes_to_group:
                shape_range.append(shape_obj.api)

            # ShapeRange 생성 및 그룹화
            grouped_shape = sheet.api.Shapes.Range([s.Name for s in shapes_to_group]).Group()
            grouped_shape.Name = group_name

            return {
                "group_created": {
                    "name": group_name,
                    "member_count": len(shapes_to_group),
                    "members": [s["name"] for s in shape_details],
                },
                "grouped_shapes": shape_details,
                "affected_shapes_count": len(shapes_to_group),
                "message": f"'{group_name}' 그룹이 성공적으로 생성되었습니다 ({len(shapes_to_group)}개 도형)",
            }

        except Exception as e:
            raise RuntimeError(f"그룹화 실패: {str(e)}")
    else:
        # macOS에서는 제한적 지원
        return {
            "grouped_shapes": shape_details,
            "group_name": group_name,
            "affected_shapes_count": len(shapes_to_group),
            "platform_limitation": "macOS에서는 그룹화 기능이 제한됩니다",
            "message": f"macOS에서는 그룹화가 제한적으로 지원됩니다",
        }


def handle_ungroup_action(sheet, target_group, all_groups, include_nested, dry_run):
    """그룹 해제 작업 처리"""
    if not target_group and not all_groups:
        raise ValueError("해제할 그룹을 지정해야 합니다 (--target-group 또는 --all-groups)")

    if target_group and all_groups:
        raise ValueError("--target-group과 --all-groups는 동시에 사용할 수 없습니다")

    groups_to_ungroup = []
    ungrouped_details = []

    # 대상 그룹 찾기
    if all_groups:
        # 모든 그룹 찾기
        for shape in sheet.shapes:
            if platform.system() == "Windows":
                try:
                    if shape.api.Type == 6:  # msoGroup
                        groups_to_ungroup.append(shape)
                except:
                    pass
            else:
                # macOS에서는 그룹 타입 확인이 제한적
                if "group" in shape.name.lower():
                    groups_to_ungroup.append(shape)
    else:
        # 특정 그룹 찾기
        target_shape = get_shape_by_name(sheet, target_group)
        if not target_shape:
            raise ValueError(f"그룹 '{target_group}'을 찾을 수 없습니다")

        # 그룹인지 확인
        if platform.system() == "Windows":
            try:
                if target_shape.api.Type != 6:  # msoGroup이 아닌 경우
                    raise ValueError(f"'{target_group}'은 그룹이 아닙니다")
            except:
                pass

        groups_to_ungroup.append(target_shape)

    if not groups_to_ungroup:
        raise ValueError("해제할 그룹을 찾을 수 없습니다")

    # 그룹 정보 수집
    for group_shape in groups_to_ungroup:
        group_info = {"group_name": group_shape.name, "member_shapes": [], "member_count": 0}

        # Windows에서 그룹 멤버 정보 수집
        if platform.system() == "Windows":
            try:
                group_items = group_shape.api.GroupItems
                group_info["member_count"] = group_items.Count

                for i in range(1, group_items.Count + 1):
                    member = group_items(i)
                    group_info["member_shapes"].append({"name": member.Name, "type": member.Type})

                # 중첩 그룹 처리
                if include_nested:
                    for i in range(1, group_items.Count + 1):
                        member = group_items(i)
                        if member.Type == 6:  # 중첩된 그룹
                            group_info["has_nested_groups"] = True

            except Exception:
                group_info["member_count"] = "알 수 없음"

        ungrouped_details.append(group_info)

    if dry_run:
        total_affected = sum(
            info.get("member_count", 0) for info in ungrouped_details if isinstance(info.get("member_count"), int)
        )

        return {
            "groups_to_ungroup": ungrouped_details,
            "total_groups": len(groups_to_ungroup),
            "affected_shapes_count": total_affected,
            "include_nested": include_nested,
            "message": f"{len(groups_to_ungroup)}개의 그룹이 해제될 예정입니다",
        }

    # 실제 그룹 해제 수행
    ungrouped_count = 0
    failed_ungroups = []

    if platform.system() == "Windows":
        for group_shape in groups_to_ungroup:
            try:
                # 중첩 그룹 처리
                if include_nested:
                    # 재귀적으로 모든 하위 그룹 해제
                    ungroup_recursively(group_shape)
                else:
                    # 단순 그룹 해제
                    group_shape.api.Ungroup()

                ungrouped_count += 1

            except Exception as e:
                failed_ungroups.append({"group_name": group_shape.name, "error": str(e)})

        total_affected = sum(
            info.get("member_count", 0) for info in ungrouped_details if isinstance(info.get("member_count"), int)
        )

        result = {
            "ungrouped_groups": ungrouped_details,
            "total_ungrouped": ungrouped_count,
            "affected_shapes_count": total_affected,
            "include_nested": include_nested,
            "message": f"{ungrouped_count}개의 그룹이 성공적으로 해제되었습니다",
        }

        if failed_ungroups:
            result["failed_ungroups"] = failed_ungroups
            result["message"] += f" ({len(failed_ungroups)}개 실패)"

        return result

    else:
        # macOS에서는 제한적 지원
        return {
            "target_groups": [g.name for g in groups_to_ungroup],
            "affected_shapes_count": len(groups_to_ungroup),
            "platform_limitation": "macOS에서는 그룹 해제 기능이 제한됩니다",
            "message": "macOS에서는 그룹 해제가 제한적으로 지원됩니다",
        }


def ungroup_recursively(group_shape):
    """중첩된 그룹을 재귀적으로 해제"""
    try:
        group_items = group_shape.api.GroupItems

        # 하위 그룹들을 먼저 해제
        nested_groups = []
        for i in range(1, group_items.Count + 1):
            member = group_items(i)
            if member.Type == 6:  # msoGroup
                nested_groups.append(member)

        # 중첩된 그룹들을 재귀적으로 해제
        for nested_group in nested_groups:
            ungroup_recursively(nested_group)

        # 최종적으로 현재 그룹 해제
        group_shape.api.Ungroup()

    except Exception:
        # 재귀 해제 실패 시 단순 해제 시도
        group_shape.api.Ungroup()


if __name__ == "__main__":
    shape_group()
