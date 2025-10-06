"""
Excel 테이블 생성 명령어 (Typer 버전)
기존 데이터 범위를 Excel Table(ListObject)로 변환
"""

import json
import platform
from pathlib import Path
from typing import Optional

import typer
import xlwings as xw

from pyhub_office_automation.version import get_version

from .engines import get_engine
from .utils import (
    ExecutionTimer,
    ExpandMode,
    create_error_response,
    create_success_response,
    get_or_open_workbook,
    get_range,
    get_sheet,
    normalize_path,
    parse_range,
    validate_range_string,
)


def table_create(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="Excel 파일의 절대 경로"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help="열린 워크북 이름으로 접근"),
    range_str: str = typer.Option(..., "--range", help="테이블로 변환할 셀 범위 (예: A1:D100, Sheet1!A1:D100)"),
    sheet: Optional[str] = typer.Option(None, "--sheet", help="시트 이름 (미지정시 활성 시트 사용)"),
    expand: Optional[ExpandMode] = typer.Option(None, "--expand", help="범위 확장 모드 (table, down, right)"),
    table_name: Optional[str] = typer.Option(None, "--table-name", help="테이블 이름 (미지정시 자동 생성)"),
    has_headers: bool = typer.Option(True, "--headers/--no-headers", help="첫 행이 헤더인지 여부"),
    table_style: str = typer.Option("TableStyleMedium2", "--table-style", help="테이블 스타일"),
    save: bool = typer.Option(True, "--save/--no-save", help="생성 후 파일 저장 여부"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택"),
    visible: bool = typer.Option(False, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부"),
):
    """
    기존 데이터 범위를 Excel Table(ListObject)로 변환합니다.

    Excel Table은 피벗테이블의 동적 범위 확장과 데이터 필터링/정렬 기능을 제공합니다.
    Windows 전용 기능으로, macOS에서는 에러가 발생합니다.

    \b
    워크북 접근 방법:
      • 옵션 없음: 활성 워크북 자동 사용 (기본값)
      • --file-path: 파일 경로로 워크북 열기
      • --workbook-name: 열린 워크북 이름으로 접근

    \b
    범위 확장 모드:
      • table: 연결된 데이터 테이블 전체로 확장
      • down: 아래쪽으로 데이터가 있는 곳까지 확장
      • right: 오른쪽으로 데이터가 있는 곳까지 확장

    \b
    테이블 스타일 예제:
      • TableStyleNone: 스타일 없음
      • TableStyleLight1~21: 밝은 테마
      • TableStyleMedium1~28: 중간 테마
      • TableStyleDark1~11: 어두운 테마

    \b
    사용 예제:
      # 기본 테이블 생성
      oa excel table-create --range "A1:D100"

      # 스타일과 이름 지정
      oa excel table-create --range "A1:D100" --table-name "SalesData" --table-style "TableStyleMedium5"

      # 자동 범위 확장
      oa excel table-create --range "A1" --expand table --table-name "AutoTable"

      # 헤더 없는 데이터
      oa excel table-create --range "A2:D100" --no-headers --table-name "RawData"

      # 특정 시트의 데이터
      oa excel table-create --range "Data!A1:F200" --table-name "DataTable"
    """
    book = None
    try:
        with ExecutionTimer() as timer:
            # 플랫폼 확인
            if platform.system() != "Windows":
                raise ValueError("Excel Table 생성은 Windows에서만 지원됩니다.")

            # 범위 문자열 유효성 검증
            if not validate_range_string(range_str):
                raise typer.BadParameter(f"잘못된 범위 형식입니다: {range_str}")

            # 워크북 연결
            book = get_or_open_workbook(file_path=file_path, workbook_name=workbook_name, visible=visible)

            # 시트 및 범위 파싱
            parsed_sheet, parsed_range = parse_range(range_str)
            sheet_name = parsed_sheet or sheet

            # 시트 가져오기
            target_sheet = get_sheet(book, sheet_name)

            # 범위 가져오기 (expand 적용)
            range_obj = get_range(target_sheet, parsed_range, expand)

            # 데이터 검증
            if not range_obj.value:
                raise ValueError("선택한 범위에 데이터가 없습니다.")

            # 테이블 이름 자동 생성
            if not table_name:
                existing_tables = [table.name for table in target_sheet.tables]
                counter = 1
                while True:
                    candidate_name = f"Table{counter}"
                    if candidate_name not in existing_tables:
                        table_name = candidate_name
                        break
                    counter += 1

            # 테이블 이름 중복 확인
            existing_table_names = [table.name for table in target_sheet.tables]
            if table_name in existing_table_names:
                raise ValueError(f"테이블 이름 '{table_name}'이 이미 존재합니다.")

            # Excel Table 생성 (Engine Layer 사용)
            try:
                # Engine 가져오기
                engine = get_engine()

                # Engine 메서드로 테이블 생성
                result = engine.create_table(
                    workbook=book.api,
                    sheet_name=target_sheet.name,
                    range_address=range_obj.address,
                    table_name=table_name,
                    has_headers=has_headers,
                    table_style=table_style,
                )

                # result는 {"name": ..., "range": ..., "row_count": ..., ...} 구조

            except Exception as e:
                raise ValueError(f"Excel Table 생성 실패: {str(e)}")

            # 저장 처리
            saved = False
            if save:
                try:
                    book.save()
                    saved = True
                except Exception as e:
                    # 저장 실패해도 테이블은 생성된 상태
                    pass

            # 생성된 테이블 정보 수집
            created_table = None
            for table in target_sheet.tables:
                if table.name == table_name:
                    created_table = table
                    break

            table_info = {
                "name": table_name,
                "range": range_obj.address,
                "sheet": target_sheet.name,
                "has_headers": has_headers,
                "style": table_style,
                "row_count": range_obj.rows.count,
                "column_count": range_obj.columns.count,
                "data_range": (
                    created_table.data_body_range.address if created_table and created_table.data_body_range else None
                ),
                "saved": saved,
            }

            # 워크북 정보 추가
            workbook_info = {
                "name": normalize_path(book.name),
                "full_name": normalize_path(book.fullname),
                "saved": getattr(book, "saved", True),
            }

            # 데이터 구성
            data_content = {
                "table": table_info,
                "workbook": workbook_info,
                "expand_mode": expand.value if expand else None,
            }

            # 성공 메시지 생성
            save_status = "저장됨" if saved else ("저장하지 않음" if not save else "저장 실패")
            message = f"Excel Table '{table_name}'을 생성했습니다 ({table_info['row_count']}행 × {table_info['column_count']}열, {save_status})"

            # 성공 응답 생성
            response = create_success_response(
                data=data_content,
                command="table-create",
                message=message,
                execution_time_ms=timer.execution_time_ms,
                book=book,
                range_obj=range_obj,
            )

            # 출력 형식에 따른 결과 반환
            if output_format == "json":
                typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
            else:  # text 형식
                table = table_info
                wb = workbook_info

                typer.echo(f"✅ {message}")
                typer.echo()
                typer.echo(f"📁 워크북: {wb['name']}")
                typer.echo(f"📄 시트: {table['sheet']}")
                typer.echo(f"🏷️ 테이블명: {table['name']}")
                typer.echo(f"📍 범위: {table['range']}")
                typer.echo(f"📊 크기: {table['row_count']}행 × {table['column_count']}열")
                typer.echo(f"🎨 스타일: {table['style']}")
                typer.echo(f"📋 헤더: {'있음' if table['has_headers'] else '없음'}")

                if saved:
                    typer.echo(f"💾 저장: ✅ 완료")
                elif not save:
                    typer.echo(f"💾 저장: ⚠️ 저장하지 않음 (--no-save 옵션)")
                else:
                    typer.echo(f"💾 저장: ❌ 실패")

    except FileNotFoundError as e:
        error_response = create_error_response(e, "table-create")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ 파일을 찾을 수 없습니다", err=True)
        raise typer.Exit(1)

    except ValueError as e:
        error_response = create_error_response(e, "table-create")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)

    except Exception as e:
        error_response = create_error_response(e, "table-create")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ 예기치 않은 오류: {str(e)}", err=True)
            typer.echo(
                "💡 Excel이 설치되어 있는지 확인하고, 파일이 다른 프로그램에서 사용 중이지 않은지 확인하세요.", err=True
            )
        raise typer.Exit(1)

    finally:
        # 워크북 정리 - 활성 워크북이나 이름으로 접근한 경우 앱 종료하지 않음
        if book is not None and not visible and file_path:
            try:
                book.app.quit()
            except:
                pass


if __name__ == "__main__":
    typer.run(table_create)
