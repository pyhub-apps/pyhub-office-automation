"""
Excel 워크북 열기 명령어
AI 에이전트와의 연동을 위한 구조화된 출력 제공
"""

import json
import sys
from pathlib import Path
import click
import xlwings as xw
from ..version import get_version


@click.command()
@click.option('--file-path', required=True,
              help='열 Excel 파일의 절대 경로')
@click.option('--visible', default=True, type=bool,
              help='Excel 애플리케이션을 화면에 표시할지 여부 (기본값: True)')
@click.option('--format', 'output_format', default='json',
              type=click.Choice(['json', 'text']),
              help='출력 형식 선택')
@click.version_option(version=get_version(), prog_name="oa excel open-workbook")
def open_workbook(file_path, visible, output_format):
    """
    Excel 워크북 파일을 엽니다.

    지정된 경로의 Excel 파일을 xlwings를 통해 열고,
    파일 정보와 시트 목록을 반환합니다.
    """
    try:
        # 파일 경로 검증
        file_path = Path(file_path).resolve()

        if not file_path.exists():
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")

        if not file_path.suffix.lower() in ['.xlsx', '.xls', '.xlsm']:
            raise ValueError(f"지원되지 않는 파일 형식입니다: {file_path.suffix}")

        # Excel 애플리케이션이 사용 가능한지 확인
        try:
            app = xw.App(visible=visible)
        except Exception as e:
            raise RuntimeError(f"Excel 애플리케이션을 시작할 수 없습니다: {str(e)}")

        # 워크북 열기
        try:
            book = app.books.open(str(file_path))
        except Exception as e:
            app.quit()
            raise RuntimeError(f"워크북을 열 수 없습니다: {str(e)}")

        # 시트 정보 수집
        sheets_info = []
        for sheet in book.sheets:
            try:
                # 시트의 사용된 범위 정보
                used_range = sheet.used_range
                if used_range:
                    last_cell = used_range.last_cell.address
                    row_count = used_range.rows.count
                    col_count = used_range.columns.count
                else:
                    last_cell = "A1"
                    row_count = 0
                    col_count = 0

                sheets_info.append({
                    "name": sheet.name,
                    "index": sheet.index,
                    "visible": sheet.visible,
                    "used_range": {
                        "last_cell": last_cell,
                        "row_count": row_count,
                        "column_count": col_count
                    }
                })
            except Exception as e:
                # 시트 정보 수집 실패 시 기본 정보만 포함
                sheets_info.append({
                    "name": sheet.name,
                    "index": sheet.index,
                    "visible": getattr(sheet, 'visible', True),
                    "error": f"시트 정보 수집 실패: {str(e)}"
                })

        # 성공 결과 데이터
        result_data = {
            "success": True,
            "command": "open-workbook",
            "version": get_version(),
            "file_info": {
                "path": str(file_path),
                "name": file_path.name,
                "size_bytes": file_path.stat().st_size,
                "exists": True
            },
            "workbook_info": {
                "name": book.name,
                "full_name": book.fullname,
                "saved": book.saved,
                "app_visible": app.visible,
                "sheet_count": len(book.sheets),
                "active_sheet": book.sheets.active.name if book.sheets else None
            },
            "sheets": sheets_info,
            "message": f"워크북이 성공적으로 열렸습니다: {file_path.name}"
        }

        # 출력 형식에 따른 결과 반환
        if output_format == 'json':
            click.echo(json.dumps(result_data, ensure_ascii=False, indent=2))
        else:
            click.echo(f"✅ 워크북 열기 성공: {file_path.name}")
            click.echo(f"📄 파일 경로: {file_path}")
            click.echo(f"📊 시트 수: {len(sheets_info)}")
            click.echo(f"🎯 활성 시트: {result_data['workbook_info']['active_sheet']}")
            if sheets_info:
                click.echo("📋 시트 목록:")
                for sheet in sheets_info:
                    if 'error' not in sheet:
                        click.echo(f"  - {sheet['name']}: {sheet['used_range']['row_count']}행 × {sheet['used_range']['column_count']}열")
                    else:
                        click.echo(f"  - {sheet['name']}: (정보 수집 실패)")

    except FileNotFoundError as e:
        error_data = {
            "success": False,
            "error_type": "FileNotFoundError",
            "error": str(e),
            "command": "open-workbook",
            "version": get_version(),
            "file_path": str(file_path)
        }

        if output_format == 'json':
            click.echo(json.dumps(error_data, ensure_ascii=False, indent=2), err=True)
        else:
            click.echo(f"❌ 파일을 찾을 수 없습니다: {file_path}", err=True)

        sys.exit(1)

    except ValueError as e:
        error_data = {
            "success": False,
            "error_type": "ValueError",
            "error": str(e),
            "command": "open-workbook",
            "version": get_version(),
            "file_path": str(file_path)
        }

        if output_format == 'json':
            click.echo(json.dumps(error_data, ensure_ascii=False, indent=2), err=True)
        else:
            click.echo(f"❌ {str(e)}", err=True)

        sys.exit(1)

    except RuntimeError as e:
        error_data = {
            "success": False,
            "error_type": "RuntimeError",
            "error": str(e),
            "command": "open-workbook",
            "version": get_version(),
            "file_path": str(file_path),
            "suggestion": "Excel이 설치되어 있는지 확인하고, 파일이 다른 프로그램에서 사용 중이지 않은지 확인하세요."
        }

        if output_format == 'json':
            click.echo(json.dumps(error_data, ensure_ascii=False, indent=2), err=True)
        else:
            click.echo(f"❌ {str(e)}", err=True)
            click.echo("💡 Excel이 설치되어 있는지 확인하고, 파일이 다른 프로그램에서 사용 중이지 않은지 확인하세요.", err=True)

        sys.exit(1)

    except Exception as e:
        error_data = {
            "success": False,
            "error_type": "UnexpectedError",
            "error": str(e),
            "command": "open-workbook",
            "version": get_version(),
            "file_path": str(file_path)
        }

        if output_format == 'json':
            click.echo(json.dumps(error_data, ensure_ascii=False, indent=2), err=True)
        else:
            click.echo(f"❌ 예기치 않은 오류: {str(e)}", err=True)

        sys.exit(1)


if __name__ == '__main__':
    open_workbook()