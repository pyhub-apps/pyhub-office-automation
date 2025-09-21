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
    create_error_response,
    create_success_response,
    format_output,
    get_or_open_workbook,
    get_range,
    get_sheet,
    get_workbook,
    normalize_path,
    parse_range,
    validate_range_string,
)


def pivot_create(
    file_path: Optional[str] = typer.Option(None, help="피벗테이블을 생성할 Excel 파일의 절대 경로"),
    use_active: bool = typer.Option(False, help="현재 활성 워크북 사용"),
    workbook_name: Optional[str] = typer.Option(None, help='열린 워크북 이름으로 접근 (예: "Sales.xlsx")'),
    source_range: str = typer.Option(..., help='소스 데이터 범위 (예: "A1:D100" 또는 "Data!A1:D100")'),
    dest_range: str = typer.Option("F1", help='피벗테이블을 생성할 위치 (기본값: "F1")'),
    dest_sheet: Optional[str] = typer.Option(None, help="피벗테이블을 생성할 시트 이름 (지정하지 않으면 현재 시트)"),
    pivot_name: Optional[str] = typer.Option(None, help="피벗테이블 이름 (지정하지 않으면 자동 생성)"),
    output_format: str = typer.Option("json", help="출력 형식 선택"),
    visible: bool = typer.Option(False, help="Excel 애플리케이션을 화면에 표시할지 여부 (기본값: False)"),
    save: bool = typer.Option(True, help="생성 후 파일 저장 여부 (기본값: True)"),
):
    """
    소스 데이터에서 피벗테이블을 생성합니다.

    기본적인 피벗테이블을 생성하며, 이후 pivot-configure 명령어로 필드 설정이 가능합니다.
    Windows 전용 기능으로, macOS에서는 에러가 발생합니다.

    워크북 접근 방법:
    - --file-path: 파일 경로로 워크북 열기
    - --use-active: 현재 활성 워크북 사용
    - --workbook-name: 열린 워크북 이름으로 접근

    예제:
        oa excel pivot-create --file-path "sales.xlsx" --source-range "A1:D100"
        oa excel pivot-create --use-active --source-range "Data!A1:F200" --dest-range "H1"
        oa excel pivot-create --workbook-name "Report.xlsx" --source-range "A1:E50" --pivot-name "SalesPivot"
    """
    book = None

    try:
        # Windows 전용 기능 확인
        if platform.system() != "Windows":
            raise RuntimeError("피벗테이블 생성은 Windows에서만 지원됩니다. macOS에서는 수동으로 피벗테이블을 생성해주세요.")

        # 소스 범위 파싱 및 검증
        source_sheet_name, source_range_part = parse_range(source_range)
        if not validate_range_string(source_range_part):
            raise ValueError(f"잘못된 소스 범위 형식입니다: {source_range}")

        # 목적지 범위 검증
        dest_sheet_name, dest_range_part = parse_range(dest_range)
        if not validate_range_string(dest_range_part):
            raise ValueError(f"잘못된 목적지 범위 형식입니다: {dest_range}")

        # 워크북 연결
        book = get_or_open_workbook(file_path=file_path, workbook_name=workbook_name, use_active=use_active, visible=visible)

        # 소스 시트 가져오기
        source_sheet = get_sheet(book, source_sheet_name)

        # 소스 데이터 범위 가져오기
        source_data_range = get_range(source_sheet, source_range_part)

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

        # 목적지 범위 가져오기
        dest_cell = get_range(target_sheet, dest_range_part)

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

            # PivotCache 생성
            pivot_cache = book.api.PivotCaches().Create(
                SourceType=PivotTableSourceType.xlDatabase, SourceData=source_data_range.api
            )

            # PivotTable 생성
            pivot_table = pivot_cache.CreatePivotTable(
                TableDestination=dest_cell.api, TableName=pivot_name, DefaultVersion=6  # Excel 2010+ 호환성
            )

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
