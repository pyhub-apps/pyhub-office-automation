"""
pandas DataFrame을 Excel 테이블로 쓰기 명령어
AI 에이전트와의 연동을 위한 구조화된 출력 제공
"""

import json
import sys
import platform
from pathlib import Path
import click
import xlwings as xw
import pandas as pd
from ..version import get_version
from .utils import (
    get_workbook, get_sheet, parse_range, get_range,
    format_output, create_error_response, create_success_response,
    validate_range_string, load_data_from_file, cleanup_temp_file
)


@click.command()
@click.option('--file-path', required=True,
              help='쓸 Excel 파일의 절대 경로')
@click.option('--data-file', required=True,
              help='DataFrame 데이터가 포함된 파일 경로 (CSV/JSON)')
@click.option('--range', 'range_str', default='A1',
              help='시작 위치 (기본값: "A1")')
@click.option('--sheet',
              help='시트 이름 (지정하지 않으면 활성 시트)')
@click.option('--include-headers', default=True, type=bool,
              help='헤더 포함 여부 (기본값: True)')
@click.option('--table-name',
              help='Excel 테이블 이름 (지정시 Excel Table 생성, Windows 전용)')
@click.option('--save', default=True, type=bool,
              help='쓰기 후 파일 저장 여부 (기본값: True)')
@click.option('--format', 'output_format', default='json',
              type=click.Choice(['json', 'text']),
              help='출력 형식 선택')
@click.option('--visible', default=False, type=bool,
              help='Excel 애플리케이션을 화면에 표시할지 여부 (기본값: False)')
@click.option('--create-sheet', default=False, type=bool,
              help='시트가 없으면 생성할지 여부 (기본값: False)')
@click.option('--replace-data', default=False, type=bool,
              help='기존 데이터를 덮어쓸지 여부 (기본값: False)')
@click.version_option(version=get_version(), prog_name="oa excel write-table")
def write_table(file_path, data_file, range_str, sheet, include_headers, table_name,
                save, output_format, visible, create_sheet, replace_data):
    """
    pandas DataFrame을 Excel 테이블로 씁니다.

    CSV 또는 JSON 파일의 데이터를 읽어서 Excel에 테이블 형태로 씁니다.
    Windows에서는 Excel Table 객체로 생성할 수 있습니다.

    지원 형식:
    - CSV 파일: 헤더가 포함된 표준 CSV
    - JSON 파일: records 형태의 JSON 배열

    예제:
        oa excel write-table --file-path "data.xlsx" --data-file "sales.csv"
        oa excel write-table --file-path "data.xlsx" --data-file "data.json" --table-name "SalesData"
        oa excel write-table --file-path "data.xlsx" --data-file "data.csv" --sheet "NewSheet" --create-sheet
    """
    book = None

    try:
        # 데이터 파일 확인
        data_path = Path(data_file)
        if not data_path.exists():
            raise FileNotFoundError(f"데이터 파일을 찾을 수 없습니다: {data_file}")

        # 데이터 로드 (pandas 사용)
        try:
            if data_path.suffix.lower() == '.csv':
                df = pd.read_csv(data_file, encoding='utf-8')
            elif data_path.suffix.lower() == '.json':
                df = pd.read_json(data_file, orient='records')
            else:
                raise ValueError(f"지원되지 않는 파일 형식입니다: {data_path.suffix}")
        except Exception as e:
            raise ValueError(f"데이터 파일 읽기 실패: {str(e)}")

        if df.empty:
            raise ValueError("데이터 파일이 비어있습니다")

        # 범위 파싱
        parsed_sheet, parsed_range = parse_range(range_str)
        start_cell = parsed_range
        sheet_name = parsed_sheet or sheet

        # 시작 셀 유효성 검증
        if not validate_range_string(start_cell):
            raise ValueError(f"잘못된 시작 셀 형식입니다: {start_cell}")

        # 워크북 열기 또는 생성
        book = get_workbook(file_path, visible=visible)

        # 시트 가져오기 또는 생성
        try:
            target_sheet = get_sheet(book, sheet_name)
        except ValueError:
            if create_sheet and sheet_name:
                target_sheet = book.sheets.add(name=sheet_name)
            else:
                raise

        # 기존 데이터 확인 및 처리
        start_range = get_range(target_sheet, start_cell)

        if not replace_data:
            # 데이터 겹침 확인
            try:
                existing_value = start_range.value
                if existing_value is not None and existing_value != "":
                    click.echo("⚠️ 경고: 시작 위치에 기존 데이터가 있습니다. --replace-data 옵션을 사용하여 덮어쓸 수 있습니다.", err=True)
            except:
                pass

        # DataFrame을 Excel 형태로 변환
        if include_headers:
            # 헤더 포함 데이터
            excel_data = [df.columns.tolist()] + df.values.tolist()
            data_rows = len(df) + 1
        else:
            # 데이터만
            excel_data = df.values.tolist()
            data_rows = len(df)

        data_cols = len(df.columns)

        # Excel에 데이터 쓰기
        try:
            start_range.value = excel_data
        except Exception as e:
            raise RuntimeError(f"데이터 쓰기 실패: {str(e)}")

        # 쓰여진 범위 계산
        try:
            end_cell = start_range.offset(data_rows - 1, data_cols - 1)
            data_range = target_sheet.range(start_range, end_cell)
            written_address = data_range.address
        except:
            written_address = start_range.address

        # Excel Table 생성 (Windows 전용)
        table_created = False
        table_error = None

        if table_name and platform.system() == "Windows":
            try:
                # Table 생성
                excel_table = target_sheet.api.ListObjects.Add(
                    SourceType=1,  # xlSrcRange
                    Source=data_range.api,
                    XlListObjectHasHeaders=1 if include_headers else 2
                )
                excel_table.Name = table_name
                table_created = True
            except Exception as e:
                table_error = str(e)
        elif table_name and platform.system() != "Windows":
            table_error = "Excel Table 생성은 Windows에서만 지원됩니다"

        # 파일 저장
        if save:
            try:
                book.save()
                saved = True
                save_error = None
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
                "rows": data_rows,
                "columns": data_cols,
                "total_cells": data_rows * data_cols
            },
            "dataframe_info": {
                "shape": list(df.shape),
                "columns": list(df.columns),
                "dtypes": {col: str(dtype) for col, dtype in df.dtypes.items()}
            },
            "table_info": {
                "table_created": table_created,
                "table_name": table_name if table_created else None,
                "has_headers": include_headers
            },
            "sheet": target_sheet.name,
            "file_info": {
                "path": str(Path(file_path).resolve()),
                "name": Path(file_path).name,
                "saved": saved
            }
        }

        if table_error:
            data_content["table_info"]["error"] = table_error
        if save_error:
            data_content["save_error"] = save_error

        # 성공 메시지 구성
        message = f"데이터를 '{written_address}' 범위에 성공적으로 작성했습니다"
        if table_created:
            message += f" (Excel Table '{table_name}' 생성됨)"
        if saved:
            message += " (파일 저장됨)"

        response = create_success_response(
            data=data_content,
            command="write-table",
            message=message
        )

        # 출력 형식에 따른 결과 반환
        if output_format == 'json':
            click.echo(json.dumps(response, ensure_ascii=False, indent=2))
        else:  # text 형식
            click.echo(f"✅ 테이블 데이터 쓰기 성공")
            click.echo(f"📄 파일: {data_content['file_info']['name']}")
            click.echo(f"📋 시트: {target_sheet.name}")
            click.echo(f"📍 범위: {written_address}")
            click.echo(f"📊 크기: {data_rows}행 × {data_cols}열 ({data_rows * data_cols}개 셀)")

            if include_headers:
                click.echo(f"🏷️ 헤더 포함: {', '.join(df.columns[:3])}{'...' if len(df.columns) > 3 else ''}")

            if table_created:
                click.echo(f"📋 Excel Table 생성: {table_name}")
            elif table_name:
                click.echo(f"⚠️ Table 생성 실패: {table_error}")

            if saved:
                click.echo("💾 파일이 저장되었습니다")
            elif save:
                click.echo(f"⚠️ 저장 실패: {save_error}")
            else:
                click.echo("📝 파일이 저장되지 않았습니다 (--save=False)")

    except FileNotFoundError as e:
        error_response = create_error_response(e, "write-table")
        if output_format == 'json':
            click.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            click.echo(f"❌ 파일을 찾을 수 없습니다: {str(e)}", err=True)
        sys.exit(1)

    except ValueError as e:
        error_response = create_error_response(e, "write-table")
        if output_format == 'json':
            click.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            click.echo(f"❌ {str(e)}", err=True)
        sys.exit(1)

    except RuntimeError as e:
        error_response = create_error_response(e, "write-table")
        if output_format == 'json':
            click.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            click.echo(f"❌ {str(e)}", err=True)
            click.echo("💡 Excel이 설치되어 있는지 확인하고, 파일이 다른 프로그램에서 사용 중이지 않은지 확인하세요.", err=True)
        sys.exit(1)

    except Exception as e:
        error_response = create_error_response(e, "write-table")
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
    write_table()