"""
Excel 셀 범위 데이터 쓰기 명령어 (Engine 기반)
"""

import json
from pathlib import Path
from typing import Optional

import typer

from pyhub_office_automation.version import get_version

from .engines import get_engine
from .utils import (
    ExecutionTimer,
    cleanup_temp_file,
    create_error_response,
    create_success_response,
    load_data_from_file,
    normalize_path,
    parse_range,
)


def range_write(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="쓸 Excel 파일의 절대 경로"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help="열린 워크북 이름으로 접근"),
    range_str: str = typer.Option(..., "--range", help="쓸 시작 셀 위치 (예: A1, Sheet1!A1)"),
    sheet: Optional[str] = typer.Option(None, "--sheet", help="시트 이름 (미지정시 활성 시트 사용)"),
    data_file: Optional[str] = typer.Option(None, "--data-file", help="쓸 데이터가 포함된 파일 경로 (JSON/CSV)"),
    data: Optional[str] = typer.Option(None, "--data", help="직접 입력할 데이터 (JSON 형식)"),
    save: bool = typer.Option(True, "--save/--no-save", help="쓰기 후 파일 저장 여부"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택"),
    visible: bool = typer.Option(False, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부"),
    create_sheet: bool = typer.Option(False, "--create-sheet", help="시트가 없으면 생성할지 여부"),
):
    """
    Excel 셀 범위에 데이터를 씁니다.

    \b
    워크북 접근 방법:
      • 옵션 없음: 활성 워크북 자동 사용
      • --file-path: 파일 경로로 워크북 열기
      • --workbook-name: 열린 워크북 이름으로 접근

    \b
    데이터 형식:
      • 단일 값: "Hello"
      • 1차원 배열: ["A", "B", "C"]
      • 2차원 배열: [["Name", "Age"], ["John", 30], ["Jane", 25]]

    \b
    사용 예제:
      oa excel range-write --range "A1" --data '["Name", "Age"]'
      oa excel range-write --file-path "data.xlsx" --range "A1" --data-file "data.json"
      oa excel range-write --range "Sheet1!A1" --data '[[1,2,3],[4,5,6]]'
    """
    temp_file_path = None

    try:
        # 데이터 입력 검증
        if not data_file and not data:
            raise ValueError("--data-file 또는 --data 중 하나를 지정해야 합니다")

        if data_file and data:
            raise ValueError("--data-file과 --data는 동시에 사용할 수 없습니다")

        # 범위 문자열 파싱
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

            # 시트 처리
            sheet_name = parsed_sheet or sheet
            if not sheet_name:
                sheet_name = wb_info["active_sheet"]

            # 시트 존재 확인 및 생성
            if sheet_name not in wb_info["sheets"]:
                if create_sheet:
                    # 시트 생성
                    sheet_name = engine.add_sheet(book, sheet_name)
                    # 워크북 정보 갱신
                    wb_info = engine.get_workbook_info(book)
                else:
                    raise ValueError(f"시트 '{sheet_name}'을 찾을 수 없습니다. 사용 가능한 시트: {wb_info['sheets']}")

            # Engine을 통해 데이터 쓰기
            engine.write_range(book, sheet_name, start_cell, write_data, include_formulas=False)

            # 데이터 크기 계산
            if isinstance(write_data, list):
                if write_data and isinstance(write_data[0], list):
                    # 2차원 데이터
                    row_count = len(write_data)
                    col_count = len(write_data[0]) if write_data else 1
                else:
                    # 1차원 데이터
                    row_count = 1
                    col_count = len(write_data)
            else:
                # 단일 값
                row_count = 1
                col_count = 1

            # 쓰여진 데이터 정보 수집
            written_info = {
                "range": start_cell,  # Engine은 실제 범위를 반환하지 않음
                "sheet": sheet_name,
                "row_count": row_count,
                "column_count": col_count,
                "cells_count": row_count * col_count,
            }

            # 저장 처리 (macOS의 경우 COM 객체가 아닐 수 있음)
            saved_successfully = False
            if save:
                try:
                    # Windows에서는 COM 객체의 Save 메서드 호출
                    if hasattr(book, "Save"):
                        book.Save()
                        saved_successfully = True
                    # macOS에서는 AppleScript로 저장 시도 (Engine 구현에 따라)
                except Exception as save_error:
                    # 저장 실패는 경고만 하고 계속 진행
                    written_info["save_warning"] = f"저장 실패: {str(save_error)}"

            # 데이터 구성
            data_content = {
                "written": written_info,
                "workbook": {"name": wb_info["name"], "path": wb_info["full_name"]},
                "saved": saved_successfully if save else False,
            }

            # 성공 메시지
            if save and saved_successfully:
                message = f"범위 '{start_cell}'에 데이터를 성공적으로 쓰고 저장했습니다"
            elif save:
                message = f"범위 '{start_cell}'에 데이터를 썼습니다 (저장 시도했으나 실패할 수 있음)"
            else:
                message = f"범위 '{start_cell}'에 데이터를 성공적으로 썼습니다"

            # 성공 응답 생성
            response = create_success_response(
                data=data_content,
                command="range-write",
                message=message,
                execution_time_ms=timer.execution_time_ms,
            )

            # 출력
            if output_format == "json":
                typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
            else:  # text 형식
                typer.echo(f"✅ {message}")
                typer.echo()
                typer.echo(f"📄 워크북: {wb_info['name']}")
                typer.echo(f"📋 시트: {sheet_name}")
                typer.echo(f"📍 범위: {start_cell}")
                typer.echo(f"📊 크기: {row_count}행 × {col_count}열 ({row_count * col_count}개 셀)")

    except FileNotFoundError as e:
        error_response = create_error_response(e, "range-write")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ 파일을 찾을 수 없습니다: {str(e)}", err=True)
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
        raise typer.Exit(1)

    finally:
        # 임시 파일 정리
        if temp_file_path:
            cleanup_temp_file(temp_file_path)


if __name__ == "__main__":
    typer.run(range_write)
