"""
Excel 워크시트 이름 변경 명령어
AI 에이전트와의 연동을 위한 구조화된 출력 제공
"""

import json
import sys
import re
from pathlib import Path
import click
import xlwings as xw
from ..version import get_version
from .utils import get_workbook, create_error_response, create_success_response


@click.command()
@click.option('--workbook', required=True,
              help='워크북 파일 경로')
@click.option('--current-name',
              help='변경할 시트의 현재 이름')
@click.option('--index', type=int,
              help='변경할 시트의 인덱스 (0부터 시작, current-name과 함께 사용 불가)')
@click.option('--new-name', required=True,
              help='새로운 시트 이름')
@click.option('--visible', default=True, type=bool,
              help='Excel 애플리케이션을 화면에 표시할지 여부')
@click.option('--format', 'output_format', default='json',
              type=click.Choice(['json', 'text']),
              help='출력 형식 선택')
@click.version_option(version=get_version(), prog_name="oa excel rename-sheet")
def rename_sheet(workbook, current_name, index, new_name, visible, output_format):
    """
    Excel 워크북의 시트 이름을 변경합니다.

    시트를 이름 또는 인덱스로 지정할 수 있습니다.
    새 이름은 워크북 내에서 고유해야 하며 Excel 시트 이름 규칙을 따라야 합니다.
    """
    try:
        # 옵션 검증
        if current_name and index is not None:
            raise ValueError("--current-name과 --index 옵션 중 하나만 지정할 수 있습니다")

        if not current_name and index is None:
            raise ValueError("--current-name 또는 --index 중 하나는 반드시 지정해야 합니다")

        # 새 이름 유효성 검증
        if not new_name or not new_name.strip():
            raise ValueError("새 시트 이름은 비어있을 수 없습니다")

        # Excel 시트 이름 규칙 검증
        invalid_chars = ['\\', '/', '*', '?', ':', '[', ']']
        for char in invalid_chars:
            if char in new_name:
                raise ValueError(f"시트 이름에 사용할 수 없는 문자가 포함되어 있습니다: '{char}'")

        if len(new_name) > 31:
            raise ValueError("시트 이름은 31자를 초과할 수 없습니다")

        # 워크북 열기
        book = get_workbook(workbook, visible=visible)

        # 기존 시트 정보 수집
        existing_sheets = [sheet.name for sheet in book.sheets]

        # 대상 시트 찾기
        target_sheet = None
        old_name = None

        if current_name:
            if current_name not in existing_sheets:
                raise ValueError(f"시트를 찾을 수 없습니다: '{current_name}'")
            target_sheet = book.sheets[current_name]
            old_name = current_name
        else:  # index 사용
            if index < 0 or index >= len(book.sheets):
                raise ValueError(f"인덱스가 범위를 벗어났습니다: {index} (0-{len(book.sheets)-1} 범위)")
            target_sheet = book.sheets[index]
            old_name = target_sheet.name

        # 새 이름 중복 검증 (자기 자신 제외)
        if new_name in existing_sheets and new_name != old_name:
            raise ValueError(f"시트 이름이 이미 존재합니다: '{new_name}'")

        # 시트 이름 변경
        try:
            target_sheet.name = new_name.strip()
        except Exception as e:
            raise RuntimeError(f"시트 이름 변경 중 오류 발생: {str(e)}")

        # 변경된 시트 정보 수집
        sheet_info = {
            "old_name": old_name,
            "new_name": target_sheet.name,
            "index": target_sheet.index,
            "visible": target_sheet.visible,
            "is_active": book.sheets.active.name == target_sheet.name
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
                "renamed_sheet": sheet_info,
                "workbook": workbook_info
            },
            command="rename-sheet",
            message=f"시트 이름이 성공적으로 변경되었습니다: '{old_name}' → '{target_sheet.name}'"
        )

        # 워크북 저장 (기존 파일 업데이트)
        try:
            book.save()
        except Exception as e:
            result_data["warning"] = f"시트 이름은 변경되었지만 저장에 실패했습니다: {str(e)}"

        # 출력 형식에 따른 결과 반환
        if output_format == 'json':
            click.echo(json.dumps(result_data, ensure_ascii=False, indent=2))
        else:
            click.echo(f"✅ 시트 이름 변경 성공")
            click.echo(f"📝 '{sheet_info['old_name']}' → '{sheet_info['new_name']}'")
            click.echo(f"📍 위치: {sheet_info['index']}번째")
            click.echo(f"🎯 활성 시트: {workbook_info['active_sheet']}")
            if result_data.get("warning"):
                click.echo(f"⚠️ {result_data['warning']}")

    except (FileNotFoundError, ValueError, RuntimeError) as e:
        error_data = create_error_response(e, "rename-sheet")

        if output_format == 'json':
            click.echo(json.dumps(error_data, ensure_ascii=False, indent=2), err=True)
        else:
            click.echo(f"❌ {str(e)}", err=True)
            if error_data.get("suggestion"):
                click.echo(f"💡 {error_data['suggestion']}", err=True)

        sys.exit(1)

    except Exception as e:
        error_data = create_error_response(e, "rename-sheet")

        if output_format == 'json':
            click.echo(json.dumps(error_data, ensure_ascii=False, indent=2), err=True)
        else:
            click.echo(f"❌ 예기치 않은 오류: {str(e)}", err=True)

        sys.exit(1)


if __name__ == '__main__':
    rename_sheet()