"""
Excel 셀 범위 데이터 쓰기 명령어 (Typer 버전)
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
    cleanup_temp_file,
    create_error_response,
    create_success_response,
    format_output,
    get_or_open_workbook,
    get_range,
    get_sheet,
    get_workbook,
    load_data_from_file,
    normalize_path,
    parse_range,
    validate_range_string,
)


def range_write(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="쓸 Excel 파일의 절대 경로"),
    use_active: bool = typer.Option(False, "--use-active", help="현재 활성 워크북 사용"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help="열린 워크북 이름으로 접근"),
    range_str: str = typer.Option(..., "--range", help="쓸 시작 셀 위치 (예: A1, Sheet1!A1)"),
    sheet: Optional[str] = typer.Option(None, "--sheet", help="시트 이름 (범위에 시트가 지정되지 않은 경우)"),
    data_file: Optional[str] = typer.Option(None, "--data-file", help="쓸 데이터가 포함된 파일 경로 (JSON/CSV)"),
    data: Optional[str] = typer.Option(None, "--data", help="직접 입력할 데이터 (JSON 형식)"),
    save: bool = typer.Option(True, "--save/--no-save", help="쓰기 후 파일 저장 여부"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택"),
    visible: bool = typer.Option(False, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부"),
    create_sheet: bool = typer.Option(False, "--create-sheet", help="시트가 없으면 생성할지 여부"),
):
    """
    Excel 셀 범위에 데이터를 씁니다.

    지정된 시작 위치부터 데이터를 쓸 수 있습니다.
    데이터는 파일에서 읽거나 직접 입력할 수 있습니다.

    워크북 접근 방법:
    - --file-path: 파일 경로로 워크북 열기 (기존 방식)
    - --use-active: 현재 활성 워크북 사용
    - --workbook-name: 열린 워크북 이름으로 접근

    데이터 형식:
    - 단일 값: "Hello"
    - 1차원 배열: ["A", "B", "C"]
    - 2차원 배열: [["Name", "Age"], ["John", 30], ["Jane", 25]]

    예제:
        oa excel range-write --file-path "data.xlsx" --range "A1" --data '["Name", "Age"]'
        oa excel range-write --use-active --range "A1" --data-file "data.json"
        oa excel range-write --workbook-name "Sales.xlsx" --range "Sheet1!A1" --data-file "data.csv"
    """
    book = None
    temp_file_path = None

    try:
        # 데이터 입력 검증
        if not data_file and not data:
            raise ValueError("--data-file 또는 --data 중 하나를 지정해야 합니다")

        if data_file and data:
            raise ValueError("--data-file과 --data는 동시에 사용할 수 없습니다")

        # 범위 문자열 유효성 검증 (시작 셀만 검증)
        parsed_sheet, parsed_range = parse_range(range_str)
        start_cell = parsed_range.split(":")[0]  # 시작 셀만 추출

        # 실행 시간 측정 시작
        with ExecutionTimer() as timer:
            # 데이터 로드
            if data_file:
                # 파일에서 데이터 읽기
                data_file_path = Path(normalize_path(data_file)).resolve()
                if not data_file_path.exists():
                    raise FileNotFoundError(f"데이터 파일을 찾을 수 없습니다: {data_file_path}")

                write_data, temp_file_path = load_data_from_file(str(data_file_path))
            else:
                # 직접 입력된 데이터 파싱
                try:
                    write_data = json.loads(data)
                except json.JSONDecodeError as e:
                    raise ValueError(f"JSON 데이터 형식이 잘못되었습니다: {str(e)}")

            # 워크북 연결
            book = get_or_open_workbook(
                file_path=file_path, workbook_name=workbook_name, use_active=use_active, visible=visible
            )

            # 시트 및 범위 처리
            sheet_name = parsed_sheet or sheet
            target_sheet = get_sheet(book, sheet_name, create_if_missing=create_sheet)

            # 시작 셀 범위 객체 가져오기
            start_range = get_range(target_sheet, start_cell)

            # 데이터의 크기를 계산하여 실제 쓸 범위 결정
            if isinstance(write_data, list):
                if len(write_data) > 0 and isinstance(write_data[0], list):
                    # 2차원 데이터
                    row_count = len(write_data)
                    col_count = len(write_data[0]) if write_data else 1
                else:
                    # 1차원 데이터 (가로로 배치)
                    row_count = 1
                    col_count = len(write_data)
            else:
                # 단일 값
                row_count = 1
                col_count = 1

            # 실제 쓸 범위 계산
            if row_count == 1 and col_count == 1:
                # 단일 셀
                write_range = start_range
                actual_range_address = start_range.address
            else:
                # 범위 확장
                end_range = start_range.offset(row_count - 1, col_count - 1)
                write_range = target_sheet.range(start_range, end_range)
                actual_range_address = write_range.address

            # 데이터 쓰기
            write_range.value = write_data

            # 쓰여진 데이터 정보 수집
            written_info = {
                "range": actual_range_address,
                "sheet": target_sheet.name,
                "data_type": type(write_data).__name__,
                "data_size": {"rows": row_count, "columns": col_count, "total_cells": row_count * col_count},
            }

            # 데이터 미리보기 추가 (큰 데이터의 경우 제한)
            if isinstance(write_data, list):
                if len(write_data) <= 5:  # 작은 데이터는 전체 포함
                    written_info["data_preview"] = write_data
                else:  # 큰 데이터는 일부만 포함
                    if isinstance(write_data[0], list):
                        written_info["data_preview"] = write_data[:3] + ["... (더 많은 데이터)"]
                    else:
                        written_info["data_preview"] = write_data[:10] + ["... (더 많은 데이터)"]
            else:
                written_info["data_preview"] = write_data

            # 저장 처리
            saved = False
            if save:
                try:
                    book.save()
                    saved = True
                except Exception as e:
                    # 저장 실패해도 데이터는 쓰여진 상태
                    written_info["save_error"] = f"저장 실패: {str(e)}"

            written_info["saved"] = saved

            # 워크북 정보 추가
            workbook_info = {
                "name": normalize_path(book.name),
                "full_name": normalize_path(book.fullname),
                "saved": getattr(book, "saved", True),
            }

            # 데이터 구성
            data_content = {
                "written": written_info,
                "workbook": workbook_info,
                "operation": {
                    "source": "data_file" if data_file else "direct_input",
                    "input_file": str(data_file_path) if data_file else None,
                },
            }

            # 성공 메시지 생성
            cells_written = row_count * col_count
            save_status = "저장됨" if saved else ("저장 실패" if save else "저장하지 않음")
            message = f"범위 '{actual_range_address}'에 {cells_written}개 셀 데이터를 썼습니다 ({save_status})"

            # 성공 응답 생성
            response = create_success_response(
                data=data_content,
                command="range-write",
                message=message,
                execution_time_ms=timer.execution_time_ms,
                book=book,
                range_obj=write_range,
                data_size=len(str(write_data).encode("utf-8")),
            )

            # 출력 형식에 따른 결과 반환
            if output_format == "json":
                typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
            else:  # text 형식
                written = written_info
                wb = workbook_info

                typer.echo(f"✅ {message}")
                typer.echo()
                typer.echo(f"📁 워크북: {wb['name']}")
                typer.echo(f"📄 시트: {written['sheet']}")
                typer.echo(f"📍 범위: {written['range']}")
                typer.echo(
                    f"📊 크기: {written['data_size']['rows']}행 × {written['data_size']['columns']}열 ({written['data_size']['total_cells']}개 셀)"
                )

                if "data_preview" in written:
                    typer.echo(f"💾 데이터 미리보기: {written['data_preview']}")

                if saved:
                    typer.echo(f"💾 저장: ✅ 완료")
                elif "save_error" in written:
                    typer.echo(f"💾 저장: ❌ {written['save_error']}")
                elif not save:
                    typer.echo(f"💾 저장: ⚠️ 저장하지 않음 (--no-save 옵션)")

    except FileNotFoundError as e:
        error_response = create_error_response(e, "range-write")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ 파일을 찾을 수 없습니다", err=True)
        raise typer.Exit(1)

    except ValueError as e:
        error_response = create_error_response(e, "range-write")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)

    except Exception as e:
        error_response = create_error_response(e, "range-write")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ 예기치 않은 오류: {str(e)}", err=True)
            typer.echo(
                "💡 Excel이 설치되어 있는지 확인하고, 파일이 다른 프로그램에서 사용 중이지 않은지 확인하세요.", err=True
            )
        raise typer.Exit(1)

    finally:
        # 임시 파일 정리
        if temp_file_path:
            cleanup_temp_file(temp_file_path)

        # 워크북 정리 - 활성 워크북이나 이름으로 접근한 경우 앱 종료하지 않음
        if book and not visible and file_path:
            try:
                book.app.quit()
            except:
                pass


if __name__ == "__main__":
    typer.run(range_write)
