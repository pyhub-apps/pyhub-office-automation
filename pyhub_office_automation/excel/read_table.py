"""
Excel 테이블 데이터를 pandas DataFrame으로 읽기 명령어
AI 에이전트와의 연동을 위한 구조화된 출력 제공
"""

import json
import sys
import tempfile
from pathlib import Path
import click
import xlwings as xw
import pandas as pd
from ..version import get_version
from .utils import (
    get_workbook, get_sheet, parse_range, get_range,
    format_output, create_error_response, create_success_response,
    validate_range_string, cleanup_temp_file
)


@click.command()
@click.option('--file-path', required=True,
              help='읽을 Excel 파일의 절대 경로')
@click.option('--range', 'range_str',
              help='읽을 테이블 범위 (지정하지 않으면 used_range 사용)')
@click.option('--sheet',
              help='시트 이름 (지정하지 않으면 활성 시트)')
@click.option('--has-headers', default=True, type=bool,
              help='첫 번째 행이 헤더인지 여부 (기본값: True)')
@click.option('--output-file',
              help='DataFrame을 저장할 파일 경로 (CSV/JSON)')
@click.option('--format', 'output_format', default='json',
              type=click.Choice(['json', 'csv', 'display']),
              help='출력 형식 선택')
@click.option('--visible', default=False, type=bool,
              help='Excel 애플리케이션을 화면에 표시할지 여부 (기본값: False)')
@click.option('--max-rows', default=None, type=int,
              help='최대 읽을 행 수 (제한 없음: None)')
@click.version_option(version=get_version(), prog_name="oa excel read-table")
def read_table(file_path, range_str, sheet, has_headers, output_file, output_format, visible, max_rows):
    """
    Excel 테이블 데이터를 pandas DataFrame으로 읽습니다.

    지정된 범위 또는 전체 사용 영역의 데이터를 테이블 형태로 읽어서
    pandas DataFrame으로 변환합니다.

    예제:
        oa excel read-table --file-path "data.xlsx"
        oa excel read-table --file-path "data.xlsx" --sheet "Sales" --output-file "sales.csv"
        oa excel read-table --file-path "data.xlsx" --range "A1:E100" --has-headers false
    """
    book = None
    temp_output_file = None

    try:
        # 워크북 열기
        book = get_workbook(file_path, visible=visible)

        # 시트 가져오기
        if range_str and '!' in range_str:
            parsed_sheet, parsed_range = parse_range(range_str)
            target_sheet = get_sheet(book, parsed_sheet)
            table_range = parsed_range
        else:
            target_sheet = get_sheet(book, sheet)
            table_range = range_str

        # 데이터 범위 결정
        if table_range:
            # 지정된 범위 사용
            if not validate_range_string(table_range):
                raise ValueError(f"잘못된 범위 형식입니다: {table_range}")
            range_obj = get_range(target_sheet, table_range)
        else:
            # used_range 사용
            range_obj = target_sheet.used_range
            if not range_obj:
                raise ValueError("시트에 데이터가 없습니다")

        # 데이터 읽기
        raw_data = range_obj.value

        if raw_data is None:
            raise ValueError("범위에 데이터가 없습니다")

        # 데이터를 DataFrame으로 변환
        if isinstance(raw_data, list):
            if len(raw_data) == 0:
                raise ValueError("범위에 데이터가 없습니다")

            # 2차원 데이터 확인
            if not isinstance(raw_data[0], list):
                # 1차원 데이터를 2차원으로 변환
                raw_data = [raw_data]

            # max_rows 적용
            if max_rows and len(raw_data) > max_rows:
                if has_headers and max_rows > 0:
                    # 헤더 + 지정된 데이터 행수
                    raw_data = raw_data[:max_rows + 1]
                else:
                    raw_data = raw_data[:max_rows]

            # DataFrame 생성
            if has_headers and len(raw_data) > 1:
                headers = raw_data[0]
                data_rows = raw_data[1:]
                df = pd.DataFrame(data_rows, columns=headers)
            else:
                df = pd.DataFrame(raw_data)

        else:
            # 단일 값
            df = pd.DataFrame([[raw_data]])

        # DataFrame 정보 수집
        df_info = {
            "shape": list(df.shape),
            "columns": list(df.columns),
            "dtypes": {col: str(dtype) for col, dtype in df.dtypes.items()},
            "non_null_counts": df.count().to_dict(),
            "memory_usage": df.memory_usage(deep=True).sum(),
            "has_headers": has_headers
        }

        # 출력 파일 저장
        saved_to_file = False
        output_file_path = None

        if output_file:
            try:
                output_path = Path(output_file)
                output_path.parent.mkdir(parents=True, exist_ok=True)

                if output_path.suffix.lower() == '.csv':
                    df.to_csv(output_path, index=False, encoding='utf-8-sig')
                elif output_path.suffix.lower() == '.json':
                    df.to_json(output_path, orient='records', force_ascii=False, indent=2)
                else:
                    # 기본적으로 CSV로 저장
                    df.to_csv(output_path, index=False, encoding='utf-8-sig')

                saved_to_file = True
                output_file_path = str(output_path.resolve())

            except Exception as e:
                # 저장 실패해도 계속 진행
                save_error = str(e)
        else:
            save_error = None

        # 응답 데이터 구성
        data_content = {
            "dataframe_info": df_info,
            "range": range_obj.address,
            "sheet": target_sheet.name,
            "file_info": {
                "path": str(Path(file_path).resolve()),
                "name": Path(file_path).name
            }
        }

        if saved_to_file:
            data_content["output_file"] = output_file_path
        elif output_file:
            data_content["save_error"] = save_error

        # 출력 형식에 따른 처리
        if output_format == 'json':
            # JSON 형태로 DataFrame 데이터 포함
            data_content["data"] = df.to_dict('records')
            response = create_success_response(
                data=data_content,
                command="read-table",
                message=f"테이블 데이터를 성공적으로 읽었습니다 ({df.shape[0]}행 × {df.shape[1]}열)"
            )
            click.echo(json.dumps(response, ensure_ascii=False, indent=2))

        elif output_format == 'csv':
            # CSV 형태로 출력
            click.echo(df.to_csv(index=False))

        else:  # display 형식
            click.echo(f"✅ 테이블 데이터 읽기 성공")
            click.echo(f"📄 파일: {data_content['file_info']['name']}")
            click.echo(f"📋 시트: {target_sheet.name}")
            click.echo(f"📍 범위: {range_obj.address}")
            click.echo(f"📊 크기: {df.shape[0]}행 × {df.shape[1]}열")

            if has_headers:
                click.echo(f"🏷️ 컬럼: {', '.join(df.columns[:5])}{'...' if len(df.columns) > 5 else ''}")

            if saved_to_file:
                click.echo(f"💾 저장됨: {output_file_path}")
            elif output_file:
                click.echo(f"⚠️ 저장 실패: {save_error}")

            # 데이터 미리보기 (상위 5행)
            if len(df) > 0:
                click.echo("\n📋 데이터 미리보기:")
                click.echo(df.head().to_string(index=False))

                if len(df) > 5:
                    click.echo(f"\n... (총 {len(df)}행)")

    except FileNotFoundError as e:
        error_response = create_error_response(e, "read-table")
        if output_format == 'json':
            click.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            click.echo(f"❌ 파일을 찾을 수 없습니다: {file_path}", err=True)
        sys.exit(1)

    except ValueError as e:
        error_response = create_error_response(e, "read-table")
        if output_format == 'json':
            click.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            click.echo(f"❌ {str(e)}", err=True)
        sys.exit(1)

    except RuntimeError as e:
        error_response = create_error_response(e, "read-table")
        if output_format == 'json':
            click.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            click.echo(f"❌ {str(e)}", err=True)
            click.echo("💡 Excel이 설치되어 있는지 확인하고, 파일이 다른 프로그램에서 사용 중이지 않은지 확인하세요.", err=True)
        sys.exit(1)

    except Exception as e:
        error_response = create_error_response(e, "read-table")
        if output_format == 'json':
            click.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            click.echo(f"❌ 예기치 않은 오류: {str(e)}", err=True)
        sys.exit(1)

    finally:
        # 임시 파일 정리
        if temp_output_file:
            cleanup_temp_file(temp_output_file)

        # 워크북 정리
        if book and not visible:
            try:
                book.app.quit()
            except:
                pass


if __name__ == '__main__':
    read_table()