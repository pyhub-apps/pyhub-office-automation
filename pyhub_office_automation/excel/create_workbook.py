"""
Excel 새 워크북 생성 명령어
AI 에이전트와의 연동을 위한 구조화된 출력 제공
"""

import json
import sys
from pathlib import Path
import click
import xlwings as xw
from ..version import get_version


@click.command()
@click.option('--name', default='NewWorkbook',
              help='생성할 워크북의 이름 (기본값: NewWorkbook)')
@click.option('--save-path',
              help='워크북을 저장할 경로 (지정하지 않으면 저장하지 않음)')
@click.option('--visible', default=True, type=bool,
              help='Excel 애플리케이션을 화면에 표시할지 여부 (기본값: True)')
@click.option('--format', 'output_format', default='json',
              type=click.Choice(['json', 'text']),
              help='출력 형식 선택')
@click.version_option(version=get_version(), prog_name="oa excel create-workbook")
def create_workbook(name, save_path, visible, output_format):
    """
    새로운 Excel 워크북을 생성합니다.

    새 워크북을 생성하고 기본 시트를 설정합니다.
    선택적으로 지정된 경로에 저장할 수 있습니다.
    """
    try:
        # Excel 애플리케이션이 사용 가능한지 확인
        try:
            app = xw.App(visible=visible)
        except Exception as e:
            raise RuntimeError(f"Excel 애플리케이션을 시작할 수 없습니다: {str(e)}")

        # 새 워크북 생성
        try:
            book = app.books.add()
        except Exception as e:
            app.quit()
            raise RuntimeError(f"새 워크북을 생성할 수 없습니다: {str(e)}")

        # 워크북 이름 설정 (저장 전까지는 임시 이름)
        original_name = book.name

        # 저장 경로가 지정된 경우 저장
        saved_path = None
        if save_path:
            try:
                save_path = Path(save_path).resolve()

                # 확장자가 없으면 .xlsx 추가
                if not save_path.suffix:
                    save_path = save_path.with_suffix('.xlsx')

                # 디렉토리가 존재하지 않으면 생성
                save_path.parent.mkdir(parents=True, exist_ok=True)

                book.save(str(save_path))
                saved_path = str(save_path)

            except Exception as e:
                # 저장 실패해도 워크북은 생성된 상태이므로 경고만 표시
                save_error = str(e)
        else:
            save_error = None

        # 시트 정보 수집
        sheets_info = []
        active_sheet = book.sheets.active if book.sheets else None
        for sheet in book.sheets:
            try:
                sheets_info.append({
                    "name": sheet.name,
                    "index": sheet.index,
                    "visible": sheet.visible,
                    "is_active": (active_sheet is not None and sheet.name == active_sheet.name)
                })
            except Exception as e:
                # 시트 정보 수집 실패 시 기본 정보만 포함
                sheets_info.append({
                    "name": getattr(sheet, 'name', 'Unknown'),
                    "index": getattr(sheet, 'index', 0),
                    "error": f"시트 정보 수집 실패: {str(e)}"
                })

        # 성공 결과 데이터
        result_data = {
            "success": True,
            "command": "create-workbook",
            "version": get_version(),
            "workbook_info": {
                "name": book.name,
                "original_name": original_name,
                "full_name": book.fullname,
                "saved": book.saved,
                "saved_path": saved_path,
                "app_visible": app.visible,
                "sheet_count": len(book.sheets),
                "active_sheet": book.sheets.active.name if book.sheets else None
            },
            "sheets": sheets_info,
            "message": f"새 워크북이 성공적으로 생성되었습니다: {book.name}"
        }

        # 저장 에러가 있는 경우 경고 추가
        if save_path and 'save_error' in locals():
            result_data["warning"] = f"워크북은 생성되었지만 저장에 실패했습니다: {save_error}"

        # 출력 형식에 따른 결과 반환
        if output_format == 'json':
            click.echo(json.dumps(result_data, ensure_ascii=False, indent=2))
        else:
            click.echo(f"✅ 새 워크북 생성 성공: {book.name}")
            if saved_path:
                click.echo(f"💾 저장 경로: {saved_path}")
            else:
                click.echo("📝 저장되지 않음 (메모리에만 존재)")
            click.echo(f"📊 시트 수: {len(sheets_info)}")
            click.echo(f"🎯 활성 시트: {result_data['workbook_info']['active_sheet']}")
            if sheets_info:
                click.echo("📋 시트 목록:")
                for sheet in sheets_info:
                    if 'error' not in sheet:
                        active_mark = " (활성)" if sheet.get('is_active') else ""
                        click.echo(f"  - {sheet['name']}{active_mark}")
                    else:
                        click.echo(f"  - (정보 수집 실패)")

            if save_path and 'save_error' in locals():
                click.echo(f"⚠️ 저장 실패: {save_error}")

    except RuntimeError as e:
        error_data = {
            "success": False,
            "error_type": "RuntimeError",
            "error": str(e),
            "command": "create-workbook",
            "version": get_version(),
            "suggestion": "Excel이 설치되어 있는지 확인하세요."
        }

        if output_format == 'json':
            click.echo(json.dumps(error_data, ensure_ascii=False, indent=2), err=True)
        else:
            click.echo(f"❌ {str(e)}", err=True)
            click.echo("💡 Excel이 설치되어 있는지 확인하세요.", err=True)

        sys.exit(1)

    except Exception as e:
        error_data = {
            "success": False,
            "error_type": "UnexpectedError",
            "error": str(e),
            "command": "create-workbook",
            "version": get_version()
        }

        if output_format == 'json':
            click.echo(json.dumps(error_data, ensure_ascii=False, indent=2), err=True)
        else:
            click.echo(f"❌ 예기치 않은 오류: {str(e)}", err=True)

        sys.exit(1)


if __name__ == '__main__':
    create_workbook()