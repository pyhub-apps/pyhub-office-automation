"""
Excel 워크시트 삭제 명령어
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
              help='삭제할 시트의 이름')
@click.option('--index', type=int,
              help='삭제할 시트의 인덱스 (0부터 시작, name과 함께 사용 불가)')
@click.option('--force', is_flag=True,
              help='확인 없이 시트 삭제 (기본값: False)')
@click.option('--visible', default=True, type=bool,
              help='Excel 애플리케이션을 화면에 표시할지 여부')
@click.option('--format', 'output_format', default='json',
              type=click.Choice(['json', 'text']),
              help='출력 형식 선택')
@click.version_option(version=get_version(), prog_name="oa excel delete-sheet")
def delete_sheet(workbook, name, index, force, visible, output_format):
    """
    Excel 워크북에서 시트를 삭제합니다.

    시트를 이름 또는 인덱스로 지정할 수 있습니다.
    마지막 시트는 삭제할 수 없으며, 워크북에 최소 1개의 시트가 유지됩니다.
    """
    try:
        # 옵션 검증
        if name and index is not None:
            raise ValueError("--name과 --index 옵션 중 하나만 지정할 수 있습니다")

        if not name and index is None:
            raise ValueError("--name 또는 --index 중 하나는 반드시 지정해야 합니다")

        # 워크북 열기
        book = get_workbook(workbook, visible=visible)

        # 최소 시트 수 확인 (워크북에는 최소 1개의 시트가 필요)
        if len(book.sheets) <= 1:
            raise ValueError("워크북에 시트가 1개만 있어서 삭제할 수 없습니다. 워크북에는 최소 1개의 시트가 필요합니다.")

        # 기존 시트 정보 수집
        existing_sheets = [sheet.name for sheet in book.sheets]
        current_active_sheet = book.sheets.active.name if book.sheets.active else None

        # 대상 시트 찾기
        target_sheet = None
        target_sheet_name = None

        if name:
            if name not in existing_sheets:
                raise ValueError(f"시트를 찾을 수 없습니다: '{name}'")
            target_sheet = book.sheets[name]
            target_sheet_name = name
        else:  # index 사용
            if index < 0 or index >= len(book.sheets):
                raise ValueError(f"인덱스가 범위를 벗어났습니다: {index} (0-{len(book.sheets)-1} 범위)")
            target_sheet = book.sheets[index]
            target_sheet_name = target_sheet.name

        # 삭제할 시트 정보 수집 (삭제 전)
        deleted_sheet_info = {
            "name": target_sheet.name,
            "index": target_sheet.index,
            "visible": target_sheet.visible,
            "was_active": current_active_sheet == target_sheet.name
        }

        # 활성 시트가 삭제 대상인 경우 다른 시트로 전환
        new_active_sheet = None
        if deleted_sheet_info["was_active"]:
            # 첫 번째 다른 시트를 활성화
            for sheet in book.sheets:
                if sheet.name != target_sheet_name:
                    try:
                        sheet.activate()
                        new_active_sheet = sheet.name
                        break
                    except Exception:
                        continue

        # 확인 메시지 (force 옵션이 없고 text 출력인 경우만)
        if not force and output_format == 'text':
            if not click.confirm(f"시트 '{target_sheet_name}'를 정말 삭제하시겠습니까?"):
                click.echo("삭제가 취소되었습니다.")
                return

        # 시트 삭제
        try:
            target_sheet.delete()
        except Exception as e:
            raise RuntimeError(f"시트 삭제 중 오류 발생: {str(e)}")

        # 삭제 후 워크북 정보 수집
        workbook_info = {
            "name": book.name,
            "full_name": book.fullname,
            "sheet_count": len(book.sheets),
            "active_sheet": book.sheets.active.name if book.sheets.active else None,
            "remaining_sheets": [sheet.name for sheet in book.sheets]
        }

        # 성공 응답 생성
        result_data = create_success_response(
            data={
                "deleted_sheet": deleted_sheet_info,
                "workbook": workbook_info,
                "new_active_sheet": new_active_sheet
            },
            command="delete-sheet",
            message=f"시트가 성공적으로 삭제되었습니다: '{target_sheet_name}'"
        )

        # 활성 시트 변경 알림
        if deleted_sheet_info["was_active"] and new_active_sheet:
            result_data["info"] = f"삭제된 시트가 활성 시트였으므로 '{new_active_sheet}' 시트가 활성화되었습니다"

        # 워크북 저장 (기존 파일 업데이트)
        try:
            book.save()
        except Exception as e:
            result_data["warning"] = f"시트는 삭제되었지만 저장에 실패했습니다: {str(e)}"

        # 출력 형식에 따른 결과 반환
        if output_format == 'json':
            click.echo(json.dumps(result_data, ensure_ascii=False, indent=2))
        else:
            click.echo(f"✅ 시트 삭제 성공: '{deleted_sheet_info['name']}'")
            click.echo(f"📊 남은 시트 수: {workbook_info['sheet_count']}")
            if deleted_sheet_info["was_active"] and new_active_sheet:
                click.echo(f"🎯 새 활성 시트: '{new_active_sheet}'")
            click.echo(f"📋 남은 시트: {', '.join(workbook_info['remaining_sheets'])}")

            if result_data.get("warning"):
                click.echo(f"⚠️ {result_data['warning']}")

    except (FileNotFoundError, ValueError, RuntimeError) as e:
        error_data = create_error_response(e, "delete-sheet")

        if output_format == 'json':
            click.echo(json.dumps(error_data, ensure_ascii=False, indent=2), err=True)
        else:
            click.echo(f"❌ {str(e)}", err=True)
            if error_data.get("suggestion"):
                click.echo(f"💡 {error_data['suggestion']}", err=True)

        sys.exit(1)

    except Exception as e:
        error_data = create_error_response(e, "delete-sheet")

        if output_format == 'json':
            click.echo(json.dumps(error_data, ensure_ascii=False, indent=2), err=True)
        else:
            click.echo(f"❌ 예기치 않은 오류: {str(e)}", err=True)

        sys.exit(1)


if __name__ == '__main__':
    delete_sheet()