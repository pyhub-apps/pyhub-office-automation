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
from .utils import get_or_open_workbook, normalize_path, ExecutionTimer, create_success_response


@click.command()
@click.option('--file-path',
              help='열 Excel 파일의 절대 경로')
@click.option('--use-active', is_flag=True,
              help='현재 활성 워크북 정보를 가져옵니다')
@click.option('--workbook-name',
              help='열린 워크북 이름으로 찾기 (예: "Sales.xlsx")')
@click.option('--visible', default=True, type=bool,
              help='Excel 애플리케이션을 화면에 표시할지 여부 (기본값: True)')
@click.option('--format', 'output_format', default='json',
              type=click.Choice(['json', 'text']),
              help='출력 형식 선택')
@click.version_option(version=get_version(), prog_name="oa excel open-workbook")
def open_workbook(file_path, use_active, workbook_name, visible, output_format):
    """
    Excel 워크북을 열거나 기존 워크북의 정보를 가져옵니다.

    다음 방법 중 하나를 사용할 수 있습니다:
    - --file-path: 지정된 경로의 파일을 엽니다
    - --use-active: 현재 활성 워크북의 정보를 가져옵니다
    - --workbook-name: 이미 열린 워크북을 이름으로 찾습니다
    """
    try:
        # 옵션 검증
        options_count = sum([bool(file_path), use_active, bool(workbook_name)])
        if options_count == 0:
            raise ValueError("--file-path, --use-active, --workbook-name 중 하나는 반드시 지정해야 합니다")
        elif options_count > 1:
            raise ValueError("--file-path, --use-active, --workbook-name 중 하나만 지정할 수 있습니다")

        # 파일 경로가 지정된 경우 파일 검증
        if file_path:
            file_path = Path(normalize_path(file_path)).resolve()
            if not file_path.exists():
                raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")
            if not file_path.suffix.lower() in ['.xlsx', '.xls', '.xlsm']:
                raise ValueError(f"지원되지 않는 파일 형식입니다: {file_path.suffix}")

        # 실행 시간 측정 시작
        with ExecutionTimer() as timer:
            # 워크북 가져오기
            book = get_or_open_workbook(
                file_path=str(file_path) if file_path else None,
                workbook_name=workbook_name,
                use_active=use_active,
                visible=visible
            )

            # 앱 객체 가져오기
            app = book.app

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

        # 응답 데이터 구성
        data_content = {
            "workbook_info": {
                "name": normalize_path(book.name),
                "full_name": normalize_path(book.fullname),
                "saved": book.saved,
                "app_visible": app.visible,
                "sheet_count": len(book.sheets),
                "active_sheet": book.sheets.active.name if book.sheets else None
            },
            "sheets": sheets_info,
            "connection_method": {
                "file_path": bool(file_path),
                "use_active": use_active,
                "workbook_name": bool(workbook_name)
            }
        }

        # 파일 정보 추가 (파일 경로가 지정된 경우에만)
        if file_path:
            data_content["file_info"] = {
                "path": str(file_path),
                "name": file_path.name,
                "size_bytes": file_path.stat().st_size,
                "exists": True
            }
            message = f"워크북이 성공적으로 열렸습니다: {file_path.name}"
        elif use_active:
            message = f"활성 워크북 정보를 가져왔습니다: {normalize_path(book.name)}"
        elif workbook_name:
            message = f"워크북을 찾았습니다: {normalize_path(book.name)}"

        # 파일 크기 계산 (통계용)
        file_size = 0
        if file_path:
            try:
                file_size = file_path.stat().st_size
            except:
                pass

        # 성공 응답 생성 (AI 에이전트 호환성 향상)
        result_data = create_success_response(
            data=data_content,
            command="workbook-open",
            message=message,
            execution_time_ms=timer.execution_time_ms,
            book=book,
            sheet_count=len(book.sheets),
            file_size=file_size
        )

        # 출력 형식에 따른 결과 반환
        if output_format == 'json':
            click.echo(json.dumps(result_data, ensure_ascii=False, indent=2))
        else:
            if use_active:
                click.echo(f"✅ 활성 워크북 정보: {normalize_path(book.name)}")
            elif workbook_name:
                click.echo(f"✅ 워크북 찾기 성공: {normalize_path(book.name)}")
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
            "file_path": str(file_path) if file_path else None
        }

        if output_format == 'json':
            click.echo(json.dumps(error_data, ensure_ascii=False, indent=2), err=True)
        else:
            click.echo(f"❌ {str(e)}", err=True)

        sys.exit(1)

    except ValueError as e:
        error_data = {
            "success": False,
            "error_type": "ValueError",
            "error": str(e),
            "command": "open-workbook",
            "version": get_version(),
            "file_path": str(file_path) if file_path else None,
            "workbook_name": workbook_name,
            "use_active": use_active
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
            "file_path": str(file_path) if file_path else None,
            "workbook_name": workbook_name,
            "use_active": use_active,
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
            "file_path": str(file_path) if file_path else None,
            "workbook_name": workbook_name,
            "use_active": use_active
        }

        if output_format == 'json':
            click.echo(json.dumps(error_data, ensure_ascii=False, indent=2), err=True)
        else:
            click.echo(f"❌ 예기치 않은 오류: {str(e)}", err=True)

        sys.exit(1)


if __name__ == '__main__':
    open_workbook()