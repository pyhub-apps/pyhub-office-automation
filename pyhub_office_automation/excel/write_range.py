"""
Excel 셀 범위 데이터 쓰기 명령어
AI 에이전트와의 연동을 위한 구조화된 출력 제공
"""

import json
import sys
from pathlib import Path
import click
import xlwings as xw
from ..version import get_version
from .utils import (
    get_workbook, get_sheet, parse_range, get_range,
    format_output, create_error_response, create_success_response,
    validate_range_string, load_data_from_file, cleanup_temp_file
)


@click.command()
@click.option('--file-path', required=True,
              help='쓸 Excel 파일의 절대 경로')
@click.option('--range', 'range_str', required=True,
              help='쓸 시작 셀 위치 (예: "A1", "Sheet1!A1")')
@click.option('--sheet',
              help='시트 이름 (범위에 시트가 지정되지 않은 경우)')
@click.option('--data-file',
              help='쓸 데이터가 포함된 파일 경로 (JSON/CSV)')
@click.option('--data',
              help='직접 입력할 데이터 (JSON 형식, 작은 데이터용)')
@click.option('--save', default=True, type=bool,
              help='쓰기 후 파일 저장 여부 (기본값: True)')
@click.option('--format', 'output_format', default='json',
              type=click.Choice(['json', 'text']),
              help='출력 형식 선택')
@click.option('--visible', default=False, type=bool,
              help='Excel 애플리케이션을 화면에 표시할지 여부 (기본값: False)')
@click.option('--create-sheet', default=False, type=bool,
              help='시트가 없으면 생성할지 여부 (기본값: False)')
@click.version_option(version=get_version(), prog_name="oa excel write-range")
def write_range(file_path, range_str, sheet, data_file, data, save, output_format, visible, create_sheet):
    """
    Excel 셀 범위에 데이터를 씁니다.

    지정된 시작 위치부터 데이터를 쓸 수 있습니다.
    데이터는 파일에서 읽거나 직접 입력할 수 있습니다.

    데이터 형식:
    - 단일 값: "Hello"
    - 1차원 배열: ["A", "B", "C"]
    - 2차원 배열: [["Name", "Age"], ["John", 30], ["Jane", 25]]

    예제:
        oa excel write-range --file-path "data.xlsx" --range "A1" --data '["Name", "Age"]'
        oa excel write-range --file-path "data.xlsx" --range "A1" --data-file "data.json"
        oa excel write-range --file-path "data.xlsx" --range "Sheet1!A1" --data-file "data.csv"
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
        start_cell = parsed_range.split(':')[0]  # 시작 셀만 추출
        if not validate_range_string(start_cell):
            raise ValueError(f"잘못된 시작 셀 형식입니다: {start_cell}")

        # 데이터 로드
        if data_file:
            write_data = load_data_from_file(data_file)
        else:
            try:
                write_data = json.loads(data)
            except json.JSONDecodeError as e:
                raise ValueError(f"데이터 JSON 파싱 오류: {str(e)}")

        # 워크북 열기 또는 생성
        book = get_workbook(file_path, visible=visible)

        # 시트 가져오기 또는 생성
        sheet_name = parsed_sheet or sheet
        try:
            target_sheet = get_sheet(book, sheet_name)
        except ValueError:
            if create_sheet and sheet_name:
                # 시트 생성
                target_sheet = book.sheets.add(name=sheet_name)
            else:
                raise

        # 시작 위치 설정
        start_range = get_range(target_sheet, start_cell)

        # 데이터 쓰기
        try:
            start_range.value = write_data
        except Exception as e:
            raise RuntimeError(f"데이터 쓰기 실패: {str(e)}")

        # 쓰여진 범위 계산
        if isinstance(write_data, list):
            if len(write_data) > 0 and isinstance(write_data[0], list):
                # 2차원 데이터
                rows = len(write_data)
                cols = len(write_data[0]) if write_data[0] else 1
            else:
                # 1차원 데이터 (가로로 배치)
                rows = 1
                cols = len(write_data)
        else:
            # 단일 값
            rows = 1
            cols = 1

        # 최종 범위 계산
        try:
            if rows == 1 and cols == 1:
                final_range = start_range
            else:
                end_cell = start_range.offset(rows - 1, cols - 1)
                final_range = target_sheet.range(start_range, end_cell)

            written_address = final_range.address
        except:
            written_address = start_range.address

        # 파일 저장
        if save:
            try:
                book.save()
                saved = True
            except Exception as e:
                saved = False
                save_error = str(e)
        else:
            saved = False
            save_error = None

        # 응답 데이터 구성
        data_content = {
            "written_range": written_address,
            "start_cell": start_range.address,
            "data_size": {
                "rows": rows,
                "columns": cols,
                "total_cells": rows * cols
            },
            "sheet": target_sheet.name,
            "file_info": {
                "path": str(Path(file_path).resolve()),
                "name": Path(file_path).name,
                "saved": saved
            }
        }

        if save_error:
            data_content["save_warning"] = f"저장 실패: {save_error}"

        # 성공 응답 생성
        message = f"데이터를 '{written_address}' 범위에 성공적으로 작성했습니다"
        if saved:
            message += " (파일 저장됨)"
        elif save:
            message += " (저장 실패)"

        response = create_success_response(
            data=data_content,
            command="write-range",
            message=message
        )

        # 출력 형식에 따른 결과 반환
        if output_format == 'json':
            click.echo(json.dumps(response, ensure_ascii=False, indent=2))
        else:  # text 형식
            click.echo(f"✅ 데이터 쓰기 성공")
            click.echo(f"📄 파일: {data_content['file_info']['name']}")
            click.echo(f"📋 시트: {target_sheet.name}")
            click.echo(f"📍 범위: {written_address}")
            click.echo(f"📊 크기: {rows}행 × {cols}열 ({rows * cols}개 셀)")

            if saved:
                click.echo("💾 파일이 저장되었습니다")
            elif save:
                click.echo(f"⚠️ 저장 실패: {save_error}")
            else:
                click.echo("📝 파일이 저장되지 않았습니다 (--save=False)")

            if data_content.get("save_warning"):
                click.echo(f"⚠️ {data_content['save_warning']}")

    except FileNotFoundError as e:
        error_response = create_error_response(e, "write-range")
        if output_format == 'json':
            click.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            click.echo(f"❌ 파일을 찾을 수 없습니다: {file_path}", err=True)
        sys.exit(1)

    except ValueError as e:
        error_response = create_error_response(e, "write-range")
        if output_format == 'json':
            click.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            click.echo(f"❌ {str(e)}", err=True)
        sys.exit(1)

    except RuntimeError as e:
        error_response = create_error_response(e, "write-range")
        if output_format == 'json':
            click.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            click.echo(f"❌ {str(e)}", err=True)
            click.echo("💡 Excel이 설치되어 있는지 확인하고, 파일이 다른 프로그램에서 사용 중이지 않은지 확인하세요.", err=True)
        sys.exit(1)

    except Exception as e:
        error_response = create_error_response(e, "write-range")
        if output_format == 'json':
            click.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            click.echo(f"❌ 예기치 않은 오류: {str(e)}", err=True)
        sys.exit(1)

    finally:
        # 임시 파일 정리
        if temp_file_path:
            cleanup_temp_file(temp_file_path)

        # 워크북 정리
        if book and not visible:
            try:
                book.app.quit()
            except:
                pass


if __name__ == '__main__':
    write_range()