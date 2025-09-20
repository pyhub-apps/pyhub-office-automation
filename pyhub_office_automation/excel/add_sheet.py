"""
Excel 워크시트 추가 명령어
AI 에이전트와의 연동을 위한 구조화된 출력 제공
"""

import json
import sys
from pathlib import Path
import click
import xlwings as xw
from ..version import get_version
from .utils import get_workbook, create_error_response, create_success_response


@click.command()
@click.option('--workbook', required=True,
              help='워크북 파일 경로')
@click.option('--name',
              help='새 시트의 이름 (지정하지 않으면 자동 생성)')
@click.option('--before',
              help='이 시트 앞에 삽입할 시트 이름')
@click.option('--after',
              help='이 시트 뒤에 삽입할 시트 이름')
@click.option('--index', type=int,
              help='삽입할 위치 인덱스 (0부터 시작)')
@click.option('--visible', default=True, type=bool,
              help='Excel 애플리케이션을 화면에 표시할지 여부')
@click.option('--format', 'output_format', default='json',
              type=click.Choice(['json', 'text']),
              help='출력 형식 선택')
@click.version_option(version=get_version(), prog_name="oa excel add-sheet")
def add_sheet(workbook, name, before, after, index, visible, output_format):
    """
    Excel 워크북에 새 워크시트를 추가합니다.

    시트 이름과 삽입 위치를 지정할 수 있습니다.
    기존 시트들 사이의 특정 위치에 삽입하거나 인덱스로 위치를 지정할 수 있습니다.
    """
    try:
        # 워크북 열기
        book = get_workbook(workbook, visible=visible)

        # 위치 지정 옵션 검증
        if sum([bool(before), bool(after), bool(index is not None)]) > 1:
            raise ValueError("--before, --after, --index 옵션 중 하나만 지정할 수 있습니다")

        # 기존 시트 정보 수집 (위치 검증용)
        existing_sheets = [sheet.name for sheet in book.sheets]

        # before/after 시트 존재 여부 확인
        before_sheet = None
        after_sheet = None

        if before:
            if before not in existing_sheets:
                raise ValueError(f"참조 시트를 찾을 수 없습니다: '{before}'")
            before_sheet = book.sheets[before]

        if after:
            if after not in existing_sheets:
                raise ValueError(f"참조 시트를 찾을 수 없습니다: '{after}'")
            after_sheet = book.sheets[after]

        # 인덱스 유효성 검증
        if index is not None:
            if index < 0 or index > len(book.sheets):
                raise ValueError(f"인덱스가 범위를 벗어났습니다: {index} (0-{len(book.sheets)} 범위)")

        # 새 시트 이름 생성 또는 검증
        if name:
            # 중복 이름 검증
            if name in existing_sheets:
                raise ValueError(f"시트 이름이 이미 존재합니다: '{name}'")
        else:
            # 자동 이름 생성 (Sheet1, Sheet2, ...)
            base_name = "Sheet"
            counter = 1
            while f"{base_name}{counter}" in existing_sheets:
                counter += 1
            name = f"{base_name}{counter}"

        # 새 시트 추가
        try:
            if index is not None:
                # 인덱스 기반 삽입 (xlwings는 1-based index 사용)
                new_sheet = book.sheets.add(name=name, before=book.sheets[index] if index < len(book.sheets) else None)
            else:
                # before/after 기반 삽입
                new_sheet = book.sheets.add(name=name, before=before_sheet, after=after_sheet)

        except Exception as e:
            raise RuntimeError(f"시트 추가 중 오류 발생: {str(e)}")

        # 시트 정보 수집
        sheet_info = {
            "name": new_sheet.name,
            "index": new_sheet.index,
            "visible": new_sheet.visible,
            "is_active": book.sheets.active.name == new_sheet.name
        }

        # 워크북 정보 업데이트
        workbook_info = {
            "name": book.name,
            "full_name": book.fullname,
            "sheet_count": len(book.sheets),
            "active_sheet": book.sheets.active.name,
            "all_sheets": [sheet.name for sheet in book.sheets]
        }

        # 성공 응답 생성
        result_data = create_success_response(
            data={
                "new_sheet": sheet_info,
                "workbook": workbook_info
            },
            command="add-sheet",
            message=f"시트가 성공적으로 추가되었습니다: '{new_sheet.name}'"
        )

        # 워크북 저장 (기존 파일 업데이트)
        try:
            book.save()
        except Exception as e:
            result_data["warning"] = f"시트는 추가되었지만 저장에 실패했습니다: {str(e)}"

        # 출력 형식에 따른 결과 반환
        if output_format == 'json':
            click.echo(json.dumps(result_data, ensure_ascii=False, indent=2))
        else:
            click.echo(f"✅ 시트 추가 성공: '{sheet_info['name']}'")
            click.echo(f"📍 위치: {sheet_info['index']}번째")
            click.echo(f"📊 전체 시트 수: {workbook_info['sheet_count']}")
            click.echo(f"🎯 활성 시트: {workbook_info['active_sheet']}")
            if result_data.get("warning"):
                click.echo(f"⚠️ {result_data['warning']}")

    except (FileNotFoundError, ValueError, RuntimeError) as e:
        error_data = create_error_response(e, "add-sheet")

        if output_format == 'json':
            click.echo(json.dumps(error_data, ensure_ascii=False, indent=2), err=True)
        else:
            click.echo(f"❌ {str(e)}", err=True)
            if error_data.get("suggestion"):
                click.echo(f"💡 {error_data['suggestion']}", err=True)

        sys.exit(1)

    except Exception as e:
        error_data = create_error_response(e, "add-sheet")

        if output_format == 'json':
            click.echo(json.dumps(error_data, ensure_ascii=False, indent=2), err=True)
        else:
            click.echo(f"❌ 예기치 않은 오류: {str(e)}", err=True)

        sys.exit(1)


if __name__ == '__main__':
    add_sheet()