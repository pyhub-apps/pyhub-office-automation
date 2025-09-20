"""
Excel 셀 범위 데이터 읽기 명령어
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
    validate_range_string
)


@click.command()
@click.option('--file-path', required=True,
              help='읽을 Excel 파일의 절대 경로')
@click.option('--range', 'range_str', required=True,
              help='읽을 셀 범위 (예: "A1:C10", "Sheet1!A1:C10")')
@click.option('--sheet',
              help='시트 이름 (범위에 시트가 지정되지 않은 경우)')
@click.option('--expand', type=click.Choice(['table', 'down', 'right']),
              help='범위 확장 모드')
@click.option('--include-formulas', default=False, type=bool,
              help='공식 포함 여부 (기본값: False)')
@click.option('--format', 'output_format', default='json',
              type=click.Choice(['json', 'csv', 'text']),
              help='출력 형식 선택')
@click.option('--visible', default=False, type=bool,
              help='Excel 애플리케이션을 화면에 표시할지 여부 (기본값: False)')
@click.version_option(version=get_version(), prog_name="oa excel read-range")
def read_range(file_path, range_str, sheet, expand, include_formulas, output_format, visible):
    """
    Excel 셀 범위의 데이터를 읽습니다.

    지정된 범위의 셀 값을 읽어서 구조화된 형태로 반환합니다.
    공식, 포맷팅된 값, 원시 값 등을 선택적으로 포함할 수 있습니다.

    예제:
        oa excel read-range --file-path "data.xlsx" --range "A1:C10"
        oa excel read-range --file-path "data.xlsx" --range "Sheet1!A1:C10" --format csv
        oa excel read-range --file-path "data.xlsx" --range "A1" --expand table
    """
    book = None
    try:
        # 범위 문자열 유효성 검증
        if not validate_range_string(range_str):
            raise ValueError(f"잘못된 범위 형식입니다: {range_str}")

        # 워크북 열기
        book = get_workbook(file_path, visible=visible)

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

            data_content = {
                "values": values,
                "formulas": formulas,
                "range": range_obj.address,
                "sheet": target_sheet.name
            }
        else:
            # 값만 읽기
            values = range_obj.value
            data_content = {
                "values": values,
                "range": range_obj.address,
                "sheet": target_sheet.name
            }

        # 범위 정보 추가
        try:
            if range_obj.count == 1:
                # 단일 셀
                data_content["range_info"] = {
                    "cells_count": 1,
                    "is_single_cell": True,
                    "row_count": 1,
                    "column_count": 1
                }
            else:
                # 다중 셀
                data_content["range_info"] = {
                    "cells_count": range_obj.count,
                    "is_single_cell": False,
                    "row_count": range_obj.rows.count,
                    "column_count": range_obj.columns.count
                }
        except:
            # 범위 정보 수집 실패시 기본값 설정
            data_content["range_info"] = {
                "cells_count": "unknown",
                "is_single_cell": False
            }

        # 파일 정보 추가
        file_info = {
            "path": str(Path(file_path).resolve()),
            "name": Path(file_path).name,
            "sheet_name": target_sheet.name
        }
        data_content["file_info"] = file_info

        # 성공 응답 생성
        response = create_success_response(
            data=data_content,
            command="read-range",
            message=f"범위 '{range_obj.address}' 데이터를 성공적으로 읽었습니다"
        )

        # 출력 형식에 따른 결과 반환
        if output_format == 'json':
            click.echo(json.dumps(response, ensure_ascii=False, indent=2))
        elif output_format == 'csv':
            # CSV 형식으로 값만 출력
            import io
            import csv

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

            click.echo(output.getvalue().rstrip())
        else:  # text 형식
            click.echo(f"📄 파일: {file_info['name']}")
            click.echo(f"📋 시트: {target_sheet.name}")
            click.echo(f"📍 범위: {range_obj.address}")

            if data_content.get("range_info", {}).get("is_single_cell"):
                click.echo(f"💾 값: {values}")
            else:
                click.echo(f"📊 데이터 크기: {data_content.get('range_info', {}).get('row_count', '?')}행 × {data_content.get('range_info', {}).get('column_count', '?')}열")
                click.echo("💾 데이터:")
                if isinstance(values, list):
                    for i, row in enumerate(values):
                        if isinstance(row, list):
                            click.echo(f"  {i+1}: {row}")
                        else:
                            click.echo(f"  {i+1}: {row}")
                else:
                    click.echo(f"  {values}")

    except FileNotFoundError as e:
        error_response = create_error_response(e, "read-range")
        if output_format == 'json':
            click.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            click.echo(f"❌ 파일을 찾을 수 없습니다: {file_path}", err=True)
        sys.exit(1)

    except ValueError as e:
        error_response = create_error_response(e, "read-range")
        if output_format == 'json':
            click.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            click.echo(f"❌ {str(e)}", err=True)
        sys.exit(1)

    except RuntimeError as e:
        error_response = create_error_response(e, "read-range")
        if output_format == 'json':
            click.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            click.echo(f"❌ {str(e)}", err=True)
            click.echo("💡 Excel이 설치되어 있는지 확인하고, 파일이 다른 프로그램에서 사용 중이지 않은지 확인하세요.", err=True)
        sys.exit(1)

    except Exception as e:
        error_response = create_error_response(e, "read-range")
        if output_format == 'json':
            click.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            click.echo(f"❌ 예기치 않은 오류: {str(e)}", err=True)
        sys.exit(1)

    finally:
        # 워크북 정리
        if book and not visible:
            try:
                book.app.quit()
            except:
                pass


if __name__ == '__main__':
    read_range()