"""
Excel 워크북 상세 정보 조회 명령어
특정 워크북의 상세 정보를 조회하여 AI 에이전트가 작업 컨텍스트를 파악할 수 있도록 지원
"""

import json
import sys
import datetime
from pathlib import Path
import click
import xlwings as xw
from ..version import get_version
from .utils import (
    get_or_open_workbook, normalize_path,
    create_success_response, create_error_response
)


@click.command()
@click.option('--file-path',
              help='조회할 Excel 파일의 절대 경로')
@click.option('--use-active', is_flag=True,
              help='현재 활성 워크북 정보를 조회합니다')
@click.option('--workbook-name',
              help='열린 워크북 이름으로 찾기 (예: "Sales.xlsx")')
@click.option('--include-sheets', is_flag=True,
              help='시트 목록 및 상세 정보 포함')
@click.option('--include-names', is_flag=True,
              help='정의된 이름(Named Ranges) 포함')
@click.option('--include-properties', is_flag=True,
              help='파일 속성 정보 포함')
@click.option('--format', 'output_format', default='json',
              type=click.Choice(['json', 'text']),
              help='출력 형식 선택')
@click.version_option(version=get_version(), prog_name="oa excel workbook-info")
def workbook_info(file_path, use_active, workbook_name, include_sheets,
                  include_names, include_properties, output_format):
    """
    특정 Excel 워크북의 상세 정보를 조회합니다.

    다음 방법 중 하나를 사용할 수 있습니다:
    - --file-path: 지정된 경로의 파일 정보를 조회합니다 (파일을 열어야 함)
    - --use-active: 현재 활성 워크북의 정보를 조회합니다
    - --workbook-name: 이미 열린 워크북을 이름으로 찾아 조회합니다
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

        # 워크북 가져오기
        book = get_or_open_workbook(
            file_path=str(file_path) if file_path else None,
            workbook_name=workbook_name,
            use_active=use_active,
            visible=True
        )

        # 기본 워크북 정보 수집
        try:
            saved_status = book.saved
        except:
            saved_status = True  # 기본값으로 저장됨으로 가정

        try:
            app_visible = book.app.visible
        except:
            app_visible = True  # 기본값으로 보임으로 가정

        workbook_data = {
            "name": normalize_path(book.name),
            "full_name": normalize_path(book.fullname),
            "saved": saved_status,
            "app_visible": app_visible,
            "sheet_count": len(book.sheets),
            "active_sheet": book.sheets.active.name if book.sheets else None
        }

        # 파일 정보 추가 (파일이 실제로 존재하는 경우)
        try:
            workbook_path = Path(book.fullname)
            if workbook_path.exists():
                file_stat = workbook_path.stat()
                workbook_data.update({
                    "file_size_bytes": file_stat.st_size,
                    "last_modified": datetime.datetime.fromtimestamp(
                        file_stat.st_mtime
                    ).isoformat()
                })
        except (OSError, AttributeError):
            # 새 워크북이거나 파일 접근 불가능한 경우
            pass

        # 응답 데이터 초기화
        response_data = {
            "workbook": workbook_data,
            "connection_method": {
                "file_path": bool(file_path),
                "use_active": use_active,
                "workbook_name": bool(workbook_name)
            }
        }

        # 시트 정보 포함
        if include_sheets:
            sheets_info = []
            for sheet in book.sheets:
                try:
                    sheet_data = {
                        "name": sheet.name,
                        "index": sheet.index,
                        "visible": sheet.visible
                    }

                    # 보호 상태 확인
                    try:
                        sheet_data["protected"] = sheet.api.ProtectContents
                    except:
                        sheet_data["protected"] = False

                    # 사용된 범위 정보
                    try:
                        used_range = sheet.used_range
                        if used_range:
                            sheet_data["used_range"] = {
                                "address": used_range.address,
                                "last_cell": used_range.last_cell.address,
                                "row_count": used_range.rows.count,
                                "column_count": used_range.columns.count,
                                "cell_count": used_range.rows.count * used_range.columns.count
                            }
                        else:
                            sheet_data["used_range"] = {
                                "address": None,
                                "last_cell": "A1",
                                "row_count": 0,
                                "column_count": 0,
                                "cell_count": 0
                            }
                    except Exception as e:
                        sheet_data["used_range_error"] = f"범위 정보 수집 실패: {str(e)}"

                    # 차트 존재 여부 확인
                    try:
                        if hasattr(sheet.api, 'ChartObjects') and sheet.api.ChartObjects().Count > 0:
                            sheet_data["has_charts"] = True
                            sheet_data["chart_count"] = sheet.api.ChartObjects().Count
                        else:
                            sheet_data["has_charts"] = False
                    except:
                        pass

                    sheets_info.append(sheet_data)

                except Exception as e:
                    # 시트 정보 수집 실패 시 기본 정보만 포함
                    sheets_info.append({
                        "name": getattr(sheet, 'name', 'Unknown'),
                        "index": getattr(sheet, 'index', -1),
                        "error": f"시트 정보 수집 실패: {str(e)}"
                    })

            response_data["sheets"] = sheets_info

        # 정의된 이름(Named Ranges) 포함
        if include_names:
            try:
                names_info = []
                for name in book.names:
                    try:
                        names_info.append({
                            "name": name.name,
                            "refers_to": name.refers_to,
                            "refers_to_range": name.refers_to_range.address if name.refers_to_range else None
                        })
                    except Exception as e:
                        names_info.append({
                            "name": getattr(name, 'name', 'Unknown'),
                            "error": f"이름 정보 수집 실패: {str(e)}"
                        })

                response_data["defined_names"] = names_info
                response_data["defined_names_count"] = len(names_info)

            except Exception as e:
                response_data["defined_names_error"] = f"정의된 이름 수집 실패: {str(e)}"

        # 파일 속성 정보 포함
        if include_properties:
            try:
                properties = {}

                # 기본 속성들
                try:
                    properties["author"] = book.api.Author
                except:
                    pass

                try:
                    properties["title"] = book.api.Title
                except:
                    pass

                try:
                    properties["subject"] = book.api.Subject
                except:
                    pass

                try:
                    properties["comments"] = book.api.Comments
                except:
                    pass

                try:
                    properties["creation_date"] = book.api.BuiltinDocumentProperties("Creation Date").Value.isoformat()
                except:
                    pass

                try:
                    properties["last_save_time"] = book.api.BuiltinDocumentProperties("Last Save Time").Value.isoformat()
                except:
                    pass

                if properties:
                    response_data["properties"] = properties

            except Exception as e:
                response_data["properties_error"] = f"속성 정보 수집 실패: {str(e)}"

        # 메시지 생성
        if use_active:
            message = f"활성 워크북 정보를 조회했습니다: {normalize_path(book.name)}"
        elif workbook_name:
            message = f"워크북 정보를 조회했습니다: {normalize_path(book.name)}"
        elif file_path:
            message = f"파일 정보를 조회했습니다: {file_path.name}"
        else:
            message = f"워크북 정보를 조회했습니다: {normalize_path(book.name)}"

        # 성공 응답 생성
        result = create_success_response(
            data=response_data,
            command="workbook-info",
            message=message
        )

        # 출력 형식에 따른 결과 반환
        if output_format == 'json':
            click.echo(json.dumps(result, ensure_ascii=False, indent=2))
        else:
            # 텍스트 형식 출력
            wb = response_data['workbook']
            click.echo(f"=== 워크북 정보: {wb['name']} ===")
            click.echo(f"파일 경로: {wb['full_name']}")
            click.echo(f"저장 상태: {'저장됨' if wb['saved'] else '저장되지 않음'}")
            click.echo(f"시트 수: {wb['sheet_count']}개")
            click.echo(f"활성 시트: {wb['active_sheet']}")

            if 'file_size_bytes' in wb:
                size_mb = wb['file_size_bytes'] / (1024 * 1024)
                click.echo(f"파일 크기: {size_mb:.1f}MB")

            if 'last_modified' in wb:
                click.echo(f"수정일: {wb['last_modified']}")

            # 시트 정보 출력
            if include_sheets and 'sheets' in response_data:
                click.echo("\n📋 시트 목록:")
                for sheet in response_data['sheets']:
                    if 'error' not in sheet:
                        protected_mark = "🔒" if sheet.get('protected', False) else ""
                        visible_mark = "👁️" if sheet.get('visible', True) else "🚫"
                        click.echo(f"  {visible_mark}{protected_mark} {sheet['name']}")

                        if 'used_range' in sheet and sheet['used_range']['address']:
                            ur = sheet['used_range']
                            click.echo(f"     사용 영역: {ur['address']} ({ur['row_count']}행 × {ur['column_count']}열)")

                        if sheet.get('has_charts'):
                            click.echo(f"     차트: {sheet.get('chart_count', 0)}개")
                    else:
                        click.echo(f"  ❌ {sheet['name']}: {sheet['error']}")

            # 정의된 이름 출력
            if include_names and 'defined_names' in response_data:
                click.echo(f"\n📌 정의된 이름: {response_data['defined_names_count']}개")
                for name_info in response_data['defined_names']:
                    if 'error' not in name_info:
                        click.echo(f"  - {name_info['name']}: {name_info.get('refers_to_range', name_info.get('refers_to', ''))}")
                    else:
                        click.echo(f"  - ❌ {name_info['name']}: {name_info['error']}")

            # 속성 정보 출력
            if include_properties and 'properties' in response_data:
                click.echo("\n📝 파일 속성:")
                props = response_data['properties']
                for key, value in props.items():
                    if value:
                        display_key = {
                            'author': '작성자',
                            'title': '제목',
                            'subject': '주제',
                            'comments': '설명',
                            'creation_date': '생성일',
                            'last_save_time': '마지막 저장'
                        }.get(key, key)
                        click.echo(f"  {display_key}: {value}")

    except FileNotFoundError as e:
        error_result = create_error_response(e, "workbook-info")

        if output_format == 'json':
            click.echo(json.dumps(error_result, ensure_ascii=False, indent=2), err=True)
        else:
            click.echo(f"❌ {str(e)}", err=True)

        sys.exit(1)

    except ValueError as e:
        error_result = create_error_response(e, "workbook-info")

        if output_format == 'json':
            click.echo(json.dumps(error_result, ensure_ascii=False, indent=2), err=True)
        else:
            click.echo(f"❌ {str(e)}", err=True)

        sys.exit(1)

    except RuntimeError as e:
        error_result = create_error_response(e, "workbook-info")
        error_result["suggestion"] = "Excel이 설치되어 있는지 확인하고, 파일이 다른 프로그램에서 사용 중이지 않은지 확인하세요."

        if output_format == 'json':
            click.echo(json.dumps(error_result, ensure_ascii=False, indent=2), err=True)
        else:
            click.echo(f"❌ {str(e)}", err=True)
            click.echo("💡 Excel이 설치되어 있는지 확인하고, 파일이 다른 프로그램에서 사용 중이지 않은지 확인하세요.", err=True)

        sys.exit(1)

    except Exception as e:
        error_result = create_error_response(e, "workbook-info")

        if output_format == 'json':
            click.echo(json.dumps(error_result, ensure_ascii=False, indent=2), err=True)
        else:
            click.echo(f"❌ 예기치 않은 오류: {str(e)}", err=True)

        sys.exit(1)


if __name__ == '__main__':
    workbook_info()