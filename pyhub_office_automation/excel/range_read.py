"""
Excel 셀 범위 데이터 읽기 명령어 (Engine 기반)
"""

import json
from pathlib import Path
from typing import Optional

import typer

from pyhub_office_automation.version import get_version

from .engines import get_engine
from .utils import (
    ExecutionTimer,
    ExpandMode,
    OutputFormat,
    create_error_response,
    create_success_response,
    normalize_path,
    parse_range,
    validate_range_string,
)


def range_read(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="읽을 Excel 파일의 절대 경로"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help="열린 워크북 이름으로 접근"),
    range_str: str = typer.Option(..., "--range", help="읽을 셀 범위 (예: A1:C10, Sheet1!A1:C10)"),
    sheet: Optional[str] = typer.Option(None, "--sheet", help="시트 이름 (미지정시 활성 시트 사용)"),
    expand: Optional[ExpandMode] = typer.Option(None, "--expand", help="범위 확장 모드 (table, down, right)"),
    include_formulas: bool = typer.Option(
        True, "--include-formulas/--no-include-formulas", help="공식 포함 여부 (기본: True)"
    ),
    output_format: OutputFormat = typer.Option(OutputFormat.JSON, "--format", help="출력 형식 선택"),
    visible: bool = typer.Option(False, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부"),
):
    """
    Excel 셀 범위의 데이터를 읽습니다.

    \b
    워크북 접근 방법:
      • 옵션 없음: 활성 워크북 자동 사용
      • --file-path: 파일 경로로 워크북 열기
      • --workbook-name: 열린 워크북 이름으로 접근

    \b
    범위 확장 모드:
      • table: 연결된 데이터 테이블 전체로 확장
      • down: 아래쪽으로 데이터가 있는 곳까지 확장
      • right: 오른쪽으로 데이터가 있는 곳까지 확장

    \b
    사용 예제:
      oa excel range-read --range "A1:C10"
      oa excel range-read --file-path "data.xlsx" --range "A1:C10"
      oa excel range-read --range "Sheet1!A1:C10" --no-include-formulas
      oa excel range-read --range "A1" --expand table
    """
    try:
        # 실행 시간 측정 시작
        with ExecutionTimer() as timer:
            # 범위 문자열 유효성 검증
            if not validate_range_string(range_str):
                raise typer.BadParameter(f"잘못된 범위 형식입니다: {range_str}")

            # Engine 획득
            engine = get_engine()

            # 워크북 가져오기
            if file_path:
                book = engine.open_workbook(file_path, visible=visible)
            elif workbook_name:
                book = engine.get_workbook_by_name(workbook_name)
            else:
                book = engine.get_active_workbook()

            # 워크북 정보 가져오기
            wb_info = engine.get_workbook_info(book)

            # 시트 및 범위 파싱
            parsed_sheet, parsed_range = parse_range(range_str)
            sheet_name = parsed_sheet or sheet

            # 시트가 지정되지 않으면 활성 시트 사용
            if not sheet_name:
                sheet_name = wb_info["active_sheet"]

            # 시트 존재 확인
            if sheet_name not in wb_info["sheets"]:
                raise ValueError(f"시트 '{sheet_name}'을 찾을 수 없습니다. 사용 가능한 시트: {wb_info['sheets']}")

            # expand 모드 문자열 변환
            expand_str = None
            if expand:
                expand_str = expand.value if hasattr(expand, "value") else str(expand)

            # Engine을 통해 범위 읽기
            range_data = engine.read_range(
                book, sheet_name, parsed_range, expand=expand_str, include_formulas=include_formulas
            )

            # 데이터 구성
            data_content = {
                "values": range_data.values,
                "range": range_data.address,
                "sheet": range_data.sheet_name,
                "range_info": {
                    "cells_count": range_data.cells_count,
                    "is_single_cell": range_data.cells_count == 1,
                    "row_count": range_data.row_count,
                    "column_count": range_data.column_count,
                },
            }

            # 공식 정보 추가 (요청된 경우)
            if include_formulas and range_data.formulas is not None:
                data_content["formulas"] = range_data.formulas

            # 파일 정보 추가
            data_content["file_info"] = {
                "path": wb_info["full_name"],
                "name": wb_info["name"],
                "sheet_name": sheet_name,
            }

            # 데이터 크기 계산
            data_size = len(str(range_data.values).encode("utf-8"))

            # 성공 응답 생성
            response = create_success_response(
                data=data_content,
                command="range-read",
                message=f"범위 '{range_data.address}' 데이터를 성공적으로 읽었습니다",
                execution_time_ms=timer.execution_time_ms,
                data_size=data_size,
            )

            # 출력 형식에 따른 결과 반환
            if output_format == OutputFormat.JSON:
                typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
            elif output_format == OutputFormat.CSV:
                # CSV 형식으로 값만 출력
                import csv
                import io

                output = io.StringIO()
                writer = csv.writer(output)

                values = range_data.values
                if isinstance(values, list):
                    if values and isinstance(values[0], list):
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
                typer.echo(f"📋 시트: {sheet_name}")
                typer.echo(f"📍 범위: {range_data.address}")

                if data_content["range_info"]["is_single_cell"]:
                    typer.echo(f"💾 값: {range_data.values}")
                else:
                    typer.echo(f"📊 데이터 크기: {range_data.row_count}행 × {range_data.column_count}열")
                    typer.echo("💾 데이터:")
                    if isinstance(range_data.values, list):
                        for i, row in enumerate(range_data.values):
                            if isinstance(row, list):
                                typer.echo(f"  {i+1}: {row}")
                            else:
                                typer.echo(f"  {i+1}: {row}")
                    else:
                        typer.echo(f"  {range_data.values}")

    except FileNotFoundError as e:
        error_response = create_error_response(e, "range-read")
        if output_format == OutputFormat.JSON:
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ 파일을 찾을 수 없습니다: {file_path}", err=True)
        raise typer.Exit(1)

    except ValueError as e:
        error_response = create_error_response(e, "range-read")
        if output_format == OutputFormat.JSON:
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)

    except Exception as e:
        error_response = create_error_response(e, "range-read")
        if output_format == OutputFormat.JSON:
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ 예기치 않은 오류: {str(e)}", err=True)
        raise typer.Exit(1)


if __name__ == "__main__":
    typer.run(range_read)
