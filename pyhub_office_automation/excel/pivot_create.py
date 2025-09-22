"""
피벗테이블 생성 명령어
Windows COM API를 활용한 Excel 피벗테이블 생성 기능
"""

import json
import platform
import sys
from pathlib import Path
from typing import Optional

import typer
import xlwings as xw

from pyhub_office_automation.version import get_version

from .utils import (
    ExpandMode,
    check_range_overlap,
    create_error_response,
    create_success_response,
    estimate_pivot_table_size,
    find_available_position,
    format_output,
    get_all_chart_ranges,
    get_all_pivot_ranges,
    get_or_open_workbook,
    get_range,
    get_sheet,
    get_workbook,
    normalize_path,
    parse_range,
    validate_auto_position_requirements,
    validate_range_string,
)


def pivot_create(
    file_path: Optional[str] = typer.Option(None, help="피벗테이블을 생성할 Excel 파일의 절대 경로"),
    workbook_name: Optional[str] = typer.Option(None, help='열린 워크북 이름으로 접근 (예: "Sales.xlsx")'),
    source_range: str = typer.Option(..., help='소스 데이터 범위 (예: "A1:D100" 또는 "Data!A1:D100")'),
    expand: Optional[ExpandMode] = typer.Option(None, "--expand", help="소스 범위 확장 모드 (table만 지원)"),
    dest_range: str = typer.Option("F1", help='피벗테이블을 생성할 위치 (기본값: "F1")'),
    dest_sheet: Optional[str] = typer.Option(None, help="피벗테이블을 생성할 시트 이름 (지정하지 않으면 현재 시트)"),
    pivot_name: Optional[str] = typer.Option(None, help="피벗테이블 이름 (지정하지 않으면 자동 생성)"),
    auto_position: bool = typer.Option(False, "--auto-position", help="자동으로 빈 공간을 찾아 배치 (Windows 전용)"),
    check_overlap: bool = typer.Option(False, "--check-overlap", help="지정된 위치의 겹침 검사 후 경고 표시"),
    spacing: int = typer.Option(2, "--spacing", help="자동 배치 시 기존 객체와의 최소 간격 (열 단위, 기본값: 2)"),
    preferred_position: str = typer.Option(
        "right", "--preferred-position", help="자동 배치 시 선호 방향 (right/bottom, 기본값: right)"
    ),
    output_format: str = typer.Option("json", help="출력 형식 선택"),
    visible: bool = typer.Option(False, help="Excel 애플리케이션을 화면에 표시할지 여부 (기본값: False)"),
    save: bool = typer.Option(True, help="생성 후 파일 저장 여부 (기본값: True)"),
):
    """
    소스 데이터에서 피벗테이블을 생성합니다.

    기본적인 피벗테이블을 생성하며, 이후 pivot-configure 명령어로 필드 설정이 가능합니다.
    Windows 전용 기능으로, macOS에서는 에러가 발생합니다.

    \b
    워크북 접근 방법:
      • 옵션 없음: 활성 워크북 자동 사용 (기본값)
      • --file-path: 파일 경로로 워크북 열기
      • --workbook-name: 열린 워크북 이름으로 접근

    \b
    소스 범위 확장 모드:
      • --expand table: 연결된 데이터 테이블 전체로 확장 (피벗테이블에 적합)
      • 범위와 expand 옵션을 함께 사용하면 시작점에서 자동으로 확장

    \b
    자동 배치 기능:
      • --auto-position: 기존 피벗테이블과 차트를 피해 자동으로 빈 공간 찾기
      • --check-overlap: 지정된 위치가 기존 객체와 겹치는지 검사
      • --spacing: 자동 배치 시 최소 간격 설정 (기본값: 2열)
      • --preferred-position: 배치 방향 선호도 (right/bottom)

    \b
    사용 예제:
      # 기본 피벗테이블 생성
      oa excel pivot-create --file-path "sales.xlsx" --source-range "A1:D100"

      # 수동 위치 지정
      oa excel pivot-create --source-range "Data!A1:F200" --dest-range "H1"

      # 자동 배치 (첫 번째 피벗 후 사용 권장)
      oa excel pivot-create --source-range "A1:D100" --auto-position

      # 자동 배치 + 사용자 설정
      oa excel pivot-create --source-range "A1:D100" --auto-position --spacing 3 --preferred-position "bottom"

      # 겹침 검사
      oa excel pivot-create --source-range "A1:D100" --dest-range "H1" --check-overlap

      # 데이터 범위 자동 확장
      oa excel pivot-create --source-range "A1" --expand table --auto-position --pivot-name "AutoPivot"
    """
    book = None

    try:
        # Windows 전용 기능 확인
        if platform.system() != "Windows":
            raise RuntimeError("피벗테이블 생성은 Windows에서만 지원됩니다. macOS에서는 수동으로 피벗테이블을 생성해주세요.")

        # expand 옵션 검증 (피벗테이블에는 table 모드만 적합)
        if expand and expand != ExpandMode.TABLE:
            raise ValueError("피벗테이블 생성에는 --expand table 옵션만 지원됩니다.")

        # 자동 배치와 수동 배치 옵션 충돌 검사
        if auto_position and dest_range != "F1":
            raise ValueError("--auto-position 옵션 사용 시 --dest-range를 지정할 수 없습니다. 자동으로 위치가 결정됩니다.")

        # preferred_position 검증
        if preferred_position not in ["right", "bottom"]:
            raise ValueError("--preferred-position은 'right' 또는 'bottom'만 지원됩니다.")

        # spacing 검증
        if spacing < 1 or spacing > 10:
            raise ValueError("--spacing은 1~10 사이의 값이어야 합니다.")

        # 소스 범위 파싱 및 검증
        source_sheet_name, source_range_part = parse_range(source_range)
        if not validate_range_string(source_range_part):
            raise ValueError(f"잘못된 소스 범위 형식입니다: {source_range}")

        # 목적지 범위 검증
        dest_sheet_name, dest_range_part = parse_range(dest_range)
        if not validate_range_string(dest_range_part):
            raise ValueError(f"잘못된 목적지 범위 형식입니다: {dest_range}")

        # 워크북 연결
        book = get_or_open_workbook(file_path=file_path, workbook_name=workbook_name, visible=visible)

        # 소스 시트 가져오기
        source_sheet = get_sheet(book, source_sheet_name)

        # 소스 데이터 범위 가져오기 (expand 옵션 적용)
        source_data_range = get_range(source_sheet, source_range_part, expand_mode=expand)

        # 소스 데이터 검증
        source_values = source_data_range.value
        if not source_values or (isinstance(source_values, list) and len(source_values) == 0):
            raise ValueError("소스 범위에 데이터가 없습니다")

        # 목적지 시트 결정
        if dest_sheet:
            try:
                target_sheet = get_sheet(book, dest_sheet)
            except ValueError:
                target_sheet = book.sheets.add(name=dest_sheet)
        else:
            target_sheet = get_sheet(book, dest_sheet_name) if dest_sheet_name else source_sheet

        # 자동 배치 또는 수동 배치 처리
        overlap_warning = None
        auto_position_info = None

        if auto_position:
            # 자동 배치 기능 사용 가능 여부 확인
            can_auto_position, auto_error = validate_auto_position_requirements(target_sheet)
            if not can_auto_position:
                raise RuntimeError(f"자동 배치를 사용할 수 없습니다: {auto_error}")

            # 피벗 테이블 예상 크기 계산
            estimated_size = estimate_pivot_table_size(source_range_part)

            # 자동으로 빈 위치 찾기
            try:
                auto_dest_range = find_available_position(
                    target_sheet, min_spacing=spacing, preferred_position=preferred_position, estimate_size=estimated_size
                )
                dest_cell = target_sheet.range(auto_dest_range)
                auto_position_info = {
                    "original_request": "auto",
                    "found_position": auto_dest_range,
                    "estimated_size": {"cols": estimated_size[0], "rows": estimated_size[1]},
                    "spacing_used": spacing,
                    "preferred_direction": preferred_position,
                }
            except RuntimeError as e:
                raise RuntimeError(f"자동 배치 실패: {str(e)}")

        else:
            # 수동 배치: 기존 로직 사용
            dest_cell = get_range(target_sheet, dest_range_part)

            # 겹침 검사 옵션 처리
            if check_overlap:
                # 피벗 테이블 예상 크기로 범위 계산
                estimated_size = estimate_pivot_table_size(source_range_part)
                dest_row = dest_cell.row
                dest_col = dest_cell.column
                estimated_end_row = dest_row + estimated_size[1] - 1
                estimated_end_col = dest_col + estimated_size[0] - 1

                from .utils import coords_to_excel_address

                estimated_range = f"{dest_cell.address}:{coords_to_excel_address(estimated_end_row, estimated_end_col)}"

                # 기존 피벗 테이블 범위 확인
                existing_pivots = get_all_pivot_ranges(target_sheet)
                overlapping_pivots = []

                for pivot_range in existing_pivots:
                    if check_range_overlap(estimated_range, pivot_range):
                        overlapping_pivots.append(pivot_range)

                # 기존 차트 범위 확인
                chart_info = get_all_chart_ranges(target_sheet)
                overlapping_charts = []

                for chart_range, _, _ in chart_info:
                    if check_range_overlap(estimated_range, chart_range):
                        overlapping_charts.append(chart_range)

                if overlapping_pivots or overlapping_charts:
                    overlap_warning = {
                        "estimated_range": estimated_range,
                        "overlapping_pivots": overlapping_pivots,
                        "overlapping_charts": overlapping_charts,
                        "recommendation": "다른 위치를 선택하거나 --auto-position 옵션을 사용하세요.",
                    }

        # 피벗테이블 이름 생성
        if not pivot_name:
            existing_pivots = []
            try:
                for pt in target_sheet.api.PivotTables():
                    existing_pivots.append(pt.Name)
            except:
                pass

            base_name = "PivotTable"
            counter = 1
            while f"{base_name}{counter}" in existing_pivots:
                counter += 1
            pivot_name = f"{base_name}{counter}"

        # Windows COM API를 사용한 피벗테이블 생성
        try:
            # xlwings constants import
            from xlwings.constants import PivotTableSourceType

            # PivotCache 생성 - 시트→부모 워크북 경로 사용 (pyhub-mcptools 방식)
            pivot_cache = source_sheet.api.Parent.PivotCaches().Create(
                SourceType=PivotTableSourceType.xlDatabase, SourceData=source_data_range.api
            )

            # PivotTable 생성 - DefaultVersion 제거, None 처리 개선
            pivot_table = pivot_cache.CreatePivotTable(TableDestination=dest_cell.api, TableName=pivot_name or None)

            # 피벗테이블 정보 수집
            pivot_info = {
                "name": pivot_table.Name,
                "source_range": source_data_range.address,
                "dest_range": dest_cell.address,
                "source_sheet": source_sheet.name,
                "dest_sheet": target_sheet.name,
                "field_count": len(source_data_range.value[0]) if isinstance(source_data_range.value, list) else 1,
                "data_rows": len(source_data_range.value) if isinstance(source_data_range.value, list) else 1,
            }

        except ImportError:
            raise RuntimeError("xlwings.constants 모듈을 가져올 수 없습니다. xlwings 최신 버전이 필요합니다.")
        except Exception as e:
            # COM 에러인 경우 더 자세한 처리를 위해 그대로 전달
            if "com_error" in str(type(e).__name__).lower():
                raise
            else:
                raise RuntimeError(f"피벗테이블 생성 실패: {str(e)}")

        # 파일 저장
        save_success = False
        save_error = None
        if save:
            try:
                book.save()
                save_success = True
            except Exception as e:
                save_error = str(e)

        # 응답 데이터 구성
        data_content = {
            "pivot_table": pivot_info,
            "source_info": {
                "range": source_data_range.address,
                "sheet": source_sheet.name,
                "data_rows": pivot_info["data_rows"],
                "field_count": pivot_info["field_count"],
            },
            "destination_info": {"range": dest_cell.address, "sheet": target_sheet.name},
            "file_info": {
                "path": (
                    str(Path(normalize_path(file_path)).resolve())
                    if file_path
                    else (normalize_path(book.fullname) if hasattr(book, "fullname") else None)
                ),
                "name": Path(normalize_path(file_path)).name if file_path else normalize_path(book.name),
                "saved": save_success,
            },
        }

        # 자동 배치 정보 추가
        if auto_position_info:
            data_content["auto_position"] = auto_position_info

        # 겹침 경고 추가
        if overlap_warning:
            data_content["overlap_warning"] = overlap_warning

        if save_error:
            data_content["save_error"] = save_error

        # 성공 메시지 구성
        message = f"피벗테이블 '{pivot_name}'이 성공적으로 생성되었습니다"
        if save_success:
            message += " (파일 저장됨)"

        response = create_success_response(data=data_content, command="pivot-create", message=message)

        # 출력 형식 검증
        if output_format not in ["json", "text"]:
            raise typer.BadParameter(f"Invalid output format: {output_format}. Must be 'json' or 'text'")

        # 출력 형식에 따른 결과 반환
        if output_format == "json":
            typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
        else:  # text 형식
            typer.echo(f"✅ 피벗테이블 생성 성공")
            typer.echo(f"📋 피벗테이블 이름: {pivot_name}")
            typer.echo(f"📄 파일: {data_content['file_info']['name']}")
            typer.echo(f"📊 소스 데이터: {source_sheet.name}!{source_data_range.address}")
            typer.echo(f"📍 생성 위치: {target_sheet.name}!{dest_cell.address}")
            typer.echo(f"📈 데이터 크기: {pivot_info['data_rows']}행 × {pivot_info['field_count']}열")

            # 자동 배치 정보 표시
            if auto_position_info:
                typer.echo(
                    f"🎯 자동 배치: {auto_position_info['found_position']} (방향: {auto_position_info['preferred_direction']}, 간격: {auto_position_info['spacing_used']}열)"
                )
                typer.echo(
                    f"📐 예상 크기: {auto_position_info['estimated_size']['cols']}열 × {auto_position_info['estimated_size']['rows']}행"
                )

            # 겹침 경고 표시
            if overlap_warning:
                typer.echo("⚠️  겹침 경고!")
                typer.echo(f"   예상 범위: {overlap_warning['estimated_range']}")
                if overlap_warning["overlapping_pivots"]:
                    typer.echo(f"   겹치는 피벗테이블: {', '.join(overlap_warning['overlapping_pivots'])}")
                if overlap_warning["overlapping_charts"]:
                    typer.echo(f"   겹치는 차트: {len(overlap_warning['overlapping_charts'])}개")
                typer.echo(f"   💡 {overlap_warning['recommendation']}")

            if save_success:
                typer.echo("💾 파일이 저장되었습니다")
            elif save:
                typer.echo(f"⚠️ 저장 실패: {save_error}")
            else:
                typer.echo("📝 파일이 저장되지 않았습니다 (--save=False)")

            typer.echo("\n💡 피벗테이블 필드 설정을 위해 'oa excel pivot-configure' 명령어를 사용하세요")

    except ValueError as e:
        error_response = create_error_response(e, "pivot-create")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)

    except RuntimeError as e:
        error_response = create_error_response(e, "pivot-create")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
            if "Windows" in str(e):
                typer.echo(
                    "💡 피벗테이블 생성은 Windows에서만 지원됩니다. macOS에서는 Excel의 수동 기능을 사용해주세요.", err=True
                )
            else:
                typer.echo("💡 Excel이 설치되어 있는지 확인하고, xlwings 최신 버전을 사용하는지 확인하세요.", err=True)
        raise typer.Exit(1)

    except Exception as e:
        error_response = create_error_response(e, "pivot-create")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ 예기치 않은 오류: {str(e)}", err=True)
        raise typer.Exit(1)

    finally:
        # 워크북 정리 - 활성 워크북이나 이름으로 접근한 경우 앱 종료하지 않음
        if book and not visible and file_path:
            try:
                book.app.quit()
            except:
                pass


if __name__ == "__main__":
    typer.run(pivot_create)
