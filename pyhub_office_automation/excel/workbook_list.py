"""
Excel 워크북 목록 조회 명령어
현재 열려있는 모든 워크북들의 목록과 기본 정보 제공
"""

import json
import sys
import datetime
from pathlib import Path
import click
import xlwings as xw
from ..version import get_version
from .utils import normalize_path, create_success_response, create_error_response, ExecutionTimer


@click.command()
@click.option('--format', 'output_format', default='json',
              type=click.Choice(['json', 'text']),
              help='출력 형식 선택')
@click.option('--detailed', is_flag=True,
              help='상세 정보 포함 (파일 경로, 시트 수, 저장 상태 등)')
@click.version_option(version=get_version(), prog_name="oa excel workbook-list")
def workbook_list(output_format, detailed):
    """
    현재 열려있는 모든 Excel 워크북 목록을 조회합니다.

    기본적으로 워크북 이름만 반환하며, --detailed 옵션으로 상세 정보를 포함할 수 있습니다.
    """
    try:
        # 실행 시간 측정 시작
        with ExecutionTimer() as timer:
            # 현재 열린 워크북들 확인
            if len(xw.books) == 0:
                # 열린 워크북이 없는 경우
                workbooks_data = []
                has_unsaved = False
                message = "현재 열려있는 워크북이 없습니다"
            else:
                workbooks_data = []
                has_unsaved = False

                for book in xw.books:
                    try:
                        # 안전하게 saved 상태 확인
                        try:
                            saved_status = book.saved
                        except:
                            saved_status = True  # 기본값으로 저장됨으로 가정

                        workbook_info = {
                            "name": normalize_path(book.name),
                            "saved": saved_status
                        }

                        # 저장되지 않은 워크북 체크
                        if not saved_status:
                            has_unsaved = True

                        if detailed:
                        # 상세 정보 추가
                        workbook_info.update({
                            "full_name": normalize_path(book.fullname),
                            "sheet_count": len(book.sheets),
                            "active_sheet": book.sheets.active.name if book.sheets else None
                        })

                        # 파일 정보 추가 (파일이 실제로 존재하는 경우)
                        try:
                            file_path = Path(book.fullname)
                            if file_path.exists():
                                file_stat = file_path.stat()
                                workbook_info.update({
                                    "file_size_bytes": file_stat.st_size,
                                    "last_modified": datetime.datetime.fromtimestamp(
                                        file_stat.st_mtime
                                    ).isoformat()
                                })
                        except (OSError, AttributeError):
                            # 새 워크북이거나 파일 접근 불가능한 경우
                            pass

                    workbooks_data.append(workbook_info)

                except Exception as e:
                    # 개별 워크북 정보 수집 실패 시 기본 정보만 포함
                    workbooks_data.append({
                        "name": getattr(book, 'name', 'Unknown'),
                        "saved": getattr(book, 'saved', False),
                        "error": f"정보 수집 실패: {str(e)}"
                    })

            # 메시지 생성
            total_count = len(workbooks_data)
            unsaved_count = len([wb for wb in workbooks_data if not wb.get('saved', True)])

            if total_count == 1:
                message = "1개의 열린 워크북을 찾았습니다"
            else:
                message = f"{total_count}개의 열린 워크북을 찾았습니다"

            if unsaved_count > 0:
                message += f" (저장되지 않음: {unsaved_count}개)"

        # 응답 데이터 구성
        response_data = {
            "workbooks": workbooks_data,
            "total_count": len(workbooks_data),
            "has_unsaved": has_unsaved
        }

        if detailed:
            # 상세 모드에서 추가 통계 정보
            saved_count = len([wb for wb in workbooks_data if wb.get('saved', True) and 'error' not in wb])
            unsaved_count = len([wb for wb in workbooks_data if not wb.get('saved', True) and 'error' not in wb])
            error_count = len([wb for wb in workbooks_data if 'error' in wb])

            response_data.update({
                "statistics": {
                    "saved_count": saved_count,
                    "unsaved_count": unsaved_count,
                    "error_count": error_count
                }
            })

        # 성공 응답 생성 (AI 에이전트 호환성 향상)
        result = create_success_response(
            data=response_data,
            command="workbook-list",
            message=message,
            execution_time_ms=timer.execution_time_ms,
            book=None,  # 특정 워크북을 대상으로 하지 않음
            workbook_count=len(workbooks_data)
        )

        # 출력 형식에 따른 결과 반환
        if output_format == 'json':
            click.echo(json.dumps(result, ensure_ascii=False, indent=2))
        else:
            # 텍스트 형식 출력
            click.echo(f"=== Excel 워크북 목록 ===")
            click.echo(f"총 {len(workbooks_data)}개의 워크북이 열려있습니다")

            if has_unsaved:
                unsaved_names = [wb['name'] for wb in workbooks_data if not wb.get('saved', True)]
                click.echo(f"⚠️  저장되지 않은 워크북: {len(unsaved_names)}개")

            click.echo()

            if not workbooks_data:
                click.echo("현재 열려있는 워크북이 없습니다.")
            else:
                for i, wb in enumerate(workbooks_data, 1):
                    saved_mark = "💾" if wb.get('saved', True) else "⚠️ "
                    click.echo(f"{i}. {saved_mark} {wb['name']}")

                    if detailed and 'error' not in wb:
                        if 'full_name' in wb:
                            click.echo(f"   경로: {wb['full_name']}")
                        if 'sheet_count' in wb:
                            click.echo(f"   시트: {wb['sheet_count']}개")
                        if 'active_sheet' in wb:
                            click.echo(f"   활성 시트: {wb['active_sheet']}")
                        if 'file_size_bytes' in wb:
                            size_mb = wb['file_size_bytes'] / (1024 * 1024)
                            click.echo(f"   크기: {size_mb:.1f}MB")
                        if 'last_modified' in wb:
                            click.echo(f"   수정일: {wb['last_modified']}")
                    elif 'error' in wb:
                        click.echo(f"   오류: {wb['error']}")

                    click.echo()

    except RuntimeError as e:
        # Excel 애플리케이션 관련 오류
        error_result = create_error_response(e, "workbook-list")
        error_result["suggestion"] = "Excel이 설치되어 있는지 확인하세요."

        if output_format == 'json':
            click.echo(json.dumps(error_result, ensure_ascii=False, indent=2), err=True)
        else:
            click.echo(f"❌ {str(e)}", err=True)
            click.echo("💡 Excel이 설치되어 있는지 확인하세요.", err=True)

        sys.exit(1)

    except Exception as e:
        # 기타 예상치 못한 오류
        error_result = create_error_response(e, "workbook-list")

        if output_format == 'json':
            click.echo(json.dumps(error_result, ensure_ascii=False, indent=2), err=True)
        else:
            click.echo(f"❌ 예기치 않은 오류: {str(e)}", err=True)

        sys.exit(1)


if __name__ == '__main__':
    workbook_list()