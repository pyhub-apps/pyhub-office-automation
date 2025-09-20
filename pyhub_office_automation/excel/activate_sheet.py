"""
Excel 워크시트 활성화 명령어
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
              help='활성화할 시트의 이름')
@click.option('--index', type=int,
              help='활성화할 시트의 인덱스 (0부터 시작, name과 함께 사용 불가)')
@click.option('--visible', default=True, type=bool,
              help='Excel 애플리케이션을 화면에 표시할지 여부')
@click.option('--format', 'output_format', default='json',
              type=click.Choice(['json', 'text']),
              help='출력 형식 선택')
@click.version_option(version=get_version(), prog_name="oa excel activate-sheet")
def activate_sheet(workbook, name, index, visible, output_format):
    """
    Excel 워크북의 특정 시트를 활성화합니다.

    시트를 이름 또는 인덱스로 지정할 수 있습니다.
    활성화된 시트는 사용자에게 표시되는 현재 시트가 됩니다.
    """
    try:
        # 옵션 검증
        if name and index is not None:
            raise ValueError("--name과 --index 옵션 중 하나만 지정할 수 있습니다")

        if not name and index is None:
            raise ValueError("--name 또는 --index 중 하나는 반드시 지정해야 합니다")

        # 워크북 열기
        book = get_workbook(workbook, visible=visible)

        # 기존 시트 정보 수집
        existing_sheets = [sheet.name for sheet in book.sheets]
        previous_active_sheet = book.sheets.active.name if book.sheets.active else None

        # 대상 시트 찾기
        target_sheet = None

        if name:
            if name not in existing_sheets:
                raise ValueError(f"시트를 찾을 수 없습니다: '{name}'")
            target_sheet = book.sheets[name]
        else:  # index 사용
            if index < 0 or index >= len(book.sheets):
                raise ValueError(f"인덱스가 범위를 벗어났습니다: {index} (0-{len(book.sheets)-1} 범위)")
            target_sheet = book.sheets[index]

        # 시트 활성화
        try:
            target_sheet.activate()
        except Exception as e:
            raise RuntimeError(f"시트 활성화 중 오류 발생: {str(e)}")

        # 활성화 후 상태 확인
        current_active_sheet = book.sheets.active.name if book.sheets.active else None
        activation_success = current_active_sheet == target_sheet.name

        # 활성화된 시트 정보 수집
        sheet_info = {
            "name": target_sheet.name,
            "index": target_sheet.index,
            "visible": target_sheet.visible,
            "is_active": activation_success,
            "previous_active_sheet": previous_active_sheet
        }

        # 워크북 정보 수집
        workbook_info = {
            "name": book.name,
            "full_name": book.fullname,
            "sheet_count": len(book.sheets),
            "active_sheet": current_active_sheet,
            "all_sheets": [
                {
                    "name": sheet.name,
                    "index": sheet.index,
                    "is_active": sheet.name == current_active_sheet
                } for sheet in book.sheets
            ]
        }

        # 성공 응답 생성
        result_data = create_success_response(
            data={
                "activated_sheet": sheet_info,
                "workbook": workbook_info
            },
            command="activate-sheet",
            message=f"시트가 성공적으로 활성화되었습니다: '{target_sheet.name}'"
        )

        # 활성화 실패 경고
        if not activation_success:
            result_data["warning"] = f"시트 활성화 명령은 실행되었지만 예상과 다른 시트가 활성화되었습니다. 현재 활성 시트: '{current_active_sheet}'"

        # 출력 형식에 따른 결과 반환
        if output_format == 'json':
            click.echo(json.dumps(result_data, ensure_ascii=False, indent=2))
        else:
            click.echo(f"✅ 시트 활성화 성공: '{sheet_info['name']}'")
            click.echo(f"📍 위치: {sheet_info['index']}번째")
            if sheet_info['previous_active_sheet'] and sheet_info['previous_active_sheet'] != sheet_info['name']:
                click.echo(f"🔄 이전 활성 시트: '{sheet_info['previous_active_sheet']}'")
            click.echo(f"📊 전체 시트 수: {workbook_info['sheet_count']}")

            if result_data.get("warning"):
                click.echo(f"⚠️ {result_data['warning']}")
            else:
                click.echo(f"🎯 현재 활성 시트: '{workbook_info['active_sheet']}'")

    except (FileNotFoundError, ValueError, RuntimeError) as e:
        error_data = create_error_response(e, "activate-sheet")

        if output_format == 'json':
            click.echo(json.dumps(error_data, ensure_ascii=False, indent=2), err=True)
        else:
            click.echo(f"❌ {str(e)}", err=True)
            if error_data.get("suggestion"):
                click.echo(f"💡 {error_data['suggestion']}", err=True)

        sys.exit(1)

    except Exception as e:
        error_data = create_error_response(e, "activate-sheet")

        if output_format == 'json':
            click.echo(json.dumps(error_data, ensure_ascii=False, indent=2), err=True)
        else:
            click.echo(f"❌ 예기치 않은 오류: {str(e)}", err=True)

        sys.exit(1)


if __name__ == '__main__':
    activate_sheet()