"""
Excel 셀 범위 데이터 읽기 명령어 (Typer 버전)
AI 에이전트와의 연동을 위한 구조화된 출력 제공
"""

import json
import sys
from pathlib import Path
from typing import Optional

import typer
import xlwings as xw

from pyhub_office_automation.version import get_version

from .utils import (
    ExecutionTimer,
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


def range_read(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="읽을 Excel 파일의 절대 경로"),
    use_active: bool = typer.Option(False, "--use-active", help="현재 활성 워크북 사용"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help="열린 워크북 이름으로 접근"),
    range_str: str = typer.Option(..., "--range", help="읽을 셀 범위 (예: A1:C10, Sheet1!A1:C10)"),
    sheet: Optional[str] = typer.Option(None, "--sheet", help="시트 이름 (범위에 시트가 지정되지 않은 경우)"),
    expand: Optional[str] = typer.Option(None, "--expand", help="범위 확장 모드"),
    include_formulas: bool = typer.Option(False, "--include-formulas", help="공식 포함 여부"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택"),
    visible: bool = typer.Option(False, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부"),
):
    """
    Excel 셀 범위의 데이터를 읽습니다.

    지정된 범위의 셀 값을 읽어서 구조화된 형태로 반환합니다.
    공식, 포맷팅된 값, 원시 값 등을 선택적으로 포함할 수 있습니다.

    워크북 접근 방법:
    - --file-path: 파일 경로로 워크북 열기 (기존 방식)
    - --use-active: 현재 활성 워크북 사용
    - --workbook-name: 열린 워크북 이름으로 접근

    예제:
        oa excel range-read --file-path "data.xlsx" --range "A1:C10"
        oa excel range-read --use-active --range "A1:C10"
        oa excel range-read --workbook-name "Sales.xlsx" --range "Sheet1!A1:C10" --format csv
        oa excel range-read --file-path "data.xlsx" --range "A1" --expand table
    """
    book = None
    try:
        # 실행 시간 측정 시작
        with ExecutionTimer() as timer:
            # 범위 문자열 유효성 검증
            if not validate_range_string(range_str):
                raise typer.BadParameter(f"잘못된 범위 형식입니다: {range_str}")

            # 워크북 연결 (새로운 통합 함수 사용)
            book = get_or_open_workbook(
                file_path=file_path, workbook_name=workbook_name, use_active=use_active, visible=visible
            )

            # 시트 및 범위 파싱
            parsed_sheet, parsed_range = parse_range(range_str)
            sheet_name = parsed_sheet or sheet

            # 시트 가져오기
            target_sheet = get_sheet(book, sheet_name)

            # 범위 가져오기
            range_obj = get_range(target_sheet, parsed_range, expand)

            # 데이터 읽기
            if include_formulas:
                # 공식과 값을 모두 읽기
                values = range_obj.value
                formulas = []

                try:
                    if range_obj.count == 1:
                        # 단일 셀인 경우
                        formulas = range_obj.formula
                    else:
                        # 다중 셀인 경우
                        formulas = range_obj.formula
                except:
                    # 공식 읽기 실패시 None으로 설정
                    formulas = None

                data_content = {"values": values, "formulas": formulas, "range": range_obj.address, "sheet": target_sheet.name}
            else:
                # 값만 읽기
                values = range_obj.value
                data_content = {"values": values, "range": range_obj.address, "sheet": target_sheet.name}

            # 범위 정보 추가
            try:
                if range_obj.count == 1:
                    # 단일 셀
                    data_content["range_info"] = {"cells_count": 1, "is_single_cell": True, "row_count": 1, "column_count": 1}
                else:
                    # 다중 셀
                    data_content["range_info"] = {
                        "cells_count": range_obj.count,
                        "is_single_cell": False,
                        "row_count": range_obj.rows.count,
                        "column_count": range_obj.columns.count,
                    }
            except:
                # 범위 정보 수집 실패시 기본값 설정
                data_content["range_info"] = {"cells_count": "unknown", "is_single_cell": False}

            # 파일 정보 추가 (file_path가 제공된 경우에만)
            if file_path:
                normalized_path = normalize_path(file_path)
                path_obj = Path(normalized_path)
                file_info = {"path": str(path_obj.resolve()), "name": path_obj.name, "sheet_name": target_sheet.name}
                data_content["file_info"] = file_info
            else:
                # 활성 워크북이나 이름으로 접근한 경우
                data_content["file_info"] = {
                    "path": normalize_path(book.fullname) if hasattr(book, "fullname") else None,
                    "name": normalize_path(book.name),
                    "sheet_name": target_sheet.name,
                }

            # 데이터 크기 계산 (통계용)
            data_size = 0
            if isinstance(values, list):
                data_size = len(str(values).encode("utf-8"))
            else:
                data_size = len(str(values).encode("utf-8"))

            # 성공 응답 생성 (AI 에이전트 호환성 향상)
            response = create_success_response(
                data=data_content,
                command="range-read",
                message=f"범위 '{range_obj.address}' 데이터를 성공적으로 읽었습니다",
                execution_time_ms=timer.execution_time_ms,
                book=book,
                range_obj=range_obj,
                data_size=data_size,
            )

            # 출력 형식에 따른 결과 반환
            if output_format == "json":
                typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
            elif output_format == "csv":
                # CSV 형식으로 값만 출력
                import csv
                import io

                output = io.StringIO()
                writer = csv.writer(output)

                if isinstance(values, list):
                    if isinstance(values[0], list):
                        # 2차원 데이터
                        writer.writerows(values)
                    else:
                        # 1차원 데이터
                        writer.writerow(values)
                else:
                    # 단일 값
                    writer.writerow([values])

                typer.echo(output.getvalue().rstrip())
            else:  # text 형식
                typer.echo(f"📄 파일: {data_content['file_info']['name']}")
                typer.echo(f"📋 시트: {target_sheet.name}")
                typer.echo(f"📍 범위: {range_obj.address}")

                if data_content.get("range_info", {}).get("is_single_cell"):
                    typer.echo(f"💾 값: {values}")
                else:
                    typer.echo(
                        f"📊 데이터 크기: {data_content.get('range_info', {}).get('row_count', '?')}행 × {data_content.get('range_info', {}).get('column_count', '?')}열"
                    )
                    typer.echo("💾 데이터:")
                    if isinstance(values, list):
                        for i, row in enumerate(values):
                            if isinstance(row, list):
                                typer.echo(f"  {i+1}: {row}")
                            else:
                                typer.echo(f"  {i+1}: {row}")
                    else:
                        typer.echo(f"  {values}")

    except FileNotFoundError as e:
        error_response = create_error_response(e, "range-read")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ 파일을 찾을 수 없습니다: {file_path}", err=True)
        raise typer.Exit(1)

    except ValueError as e:
        error_response = create_error_response(e, "range-read")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)

    except RuntimeError as e:
        error_response = create_error_response(e, "range-read")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
            typer.echo(
                "💡 Excel이 설치되어 있는지 확인하고, 파일이 다른 프로그램에서 사용 중이지 않은지 확인하세요.", err=True
            )
        raise typer.Exit(1)

    except Exception as e:
        error_response = create_error_response(e, "range-read")
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
