"""
Excel 새 워크북 생성 명령어 (Typer 버전)
AI 에이전트와의 연동을 위한 구조화된 출력 제공
"""

import json
import sys
from pathlib import Path
from typing import Optional

import typer
import xlwings as xw

from pyhub_office_automation.version import get_version

from .utils import ExecutionTimer, create_error_response, create_success_response, get_active_app, normalize_path


def workbook_create(
    name: str = typer.Option("NewWorkbook", "--name", help="생성할 워크북의 이름"),
    save_path: Optional[str] = typer.Option(None, "--save-path", help="워크북을 저장할 경로"),
    use_active: bool = typer.Option(False, "--use-active", help="기존 Excel 애플리케이션을 사용하여 새 워크북 생성"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help="특정 워크북의 애플리케이션을 사용"),
    visible: bool = typer.Option(True, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택"),
):
    """
    새로운 Excel 워크북을 생성합니다.

    항상 새로운 워크북을 생성하며, Excel 애플리케이션 연결 방식을 선택할 수 있습니다:
    - 기본: 새 Excel 애플리케이션 인스턴스 사용
    - --use-active: 현재 활성 Excel 애플리케이션 사용
    - --workbook-name: 특정 워크북의 애플리케이션 사용

    예제:
        oa excel workbook-create --name "MyReport"
        oa excel workbook-create --name "Data" --save-path "data.xlsx"
        oa excel workbook-create --use-active --name "NewSheet"
    """
    app = None
    book = None
    try:
        # 실행 시간 측정 시작
        with ExecutionTimer() as timer:
            # Excel 애플리케이션 가져오기
            if use_active:
                # 기존 활성 애플리케이션 사용
                app = get_active_app(visible=visible)
            elif workbook_name:
                # 특정 워크북의 애플리케이션 사용
                target_book = None
                for book_iter in xw.books:
                    if (
                        book_iter.name == workbook_name
                        or Path(book_iter.name).name == workbook_name
                        or Path(book_iter.name).stem == Path(workbook_name).stem
                    ):
                        target_book = book_iter
                        break

                if target_book is None:
                    raise RuntimeError(f"워크북 '{workbook_name}'을 찾을 수 없습니다")

                app = target_book.app
            else:
                # 새 Excel 애플리케이션 생성
                try:
                    app = xw.App(visible=visible)
                except Exception as e:
                    raise RuntimeError(f"Excel 애플리케이션을 시작할 수 없습니다: {str(e)}")

            # 새 워크북 생성
            try:
                book = app.books.add()
            except Exception as e:
                # 기존 앱을 사용하는 경우에는 종료하지 않음
                if not use_active and not workbook_name:
                    app.quit()
                raise RuntimeError(f"새 워크북을 생성할 수 없습니다: {str(e)}")

            # 워크북 이름 설정 (저장 전까지는 임시 이름)
            original_name = book.name

            # 저장 경로가 지정된 경우 저장
            saved_path = None
            if save_path:
                try:
                    # 경로 정규화
                    save_path_obj = Path(normalize_path(save_path)).resolve()

                    # 확장자가 없으면 .xlsx 추가
                    if not save_path_obj.suffix:
                        save_path_obj = save_path_obj.with_suffix(".xlsx")

                    # 디렉토리 생성 (필요한 경우)
                    save_path_obj.parent.mkdir(parents=True, exist_ok=True)

                    # 워크북 저장
                    book.save(str(save_path_obj))
                    saved_path = str(save_path_obj)

                except Exception as e:
                    # 저장 실패 시에도 워크북은 생성된 상태
                    raise RuntimeError(f"워크북 저장에 실패했습니다: {str(e)}")

            # 시트 정보 수집
            sheets_info = []
            for sheet in book.sheets:
                try:
                    sheet_info = {"name": sheet.name, "index": sheet.index, "is_active": sheet == book.sheets.active}
                    sheets_info.append(sheet_info)
                except Exception as e:
                    sheets_info.append({"name": getattr(sheet, "name", "Unknown"), "error": f"시트 정보 수집 실패: {str(e)}"})

            # 워크북 정보 구성
            workbook_info = {
                "name": normalize_path(book.name),
                "original_name": original_name,
                "full_name": normalize_path(book.fullname),
                "saved": getattr(book, "saved", False),
                "saved_path": saved_path,
                "sheet_count": len(book.sheets),
                "active_sheet": book.sheets.active.name if book.sheets.active else None,
                "sheets": sheets_info,
            }

            # 애플리케이션 정보
            app_info = {
                "version": getattr(app, "version", "Unknown"),
                "visible": getattr(app, "visible", visible),
                "is_new_instance": not use_active and not workbook_name,
            }

            # 데이터 구성
            data_content = {
                "workbook": workbook_info,
                "application": app_info,
                "creation_method": "active_app" if use_active else ("existing_app" if workbook_name else "new_app"),
            }

            # 성공 메시지
            if saved_path:
                message = f"새 워크북 '{workbook_info['name']}'을(를) 생성하고 '{saved_path}'에 저장했습니다"
            else:
                message = f"새 워크북 '{workbook_info['name']}'을(를) 생성했습니다"

            # 성공 응답 생성
            response = create_success_response(
                data=data_content,
                command="workbook-create",
                message=message,
                execution_time_ms=timer.execution_time_ms,
                book=book,
            )

            # 출력 형식에 따른 결과 반환
            if output_format == "json":
                typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
            else:  # text 형식
                wb = workbook_info
                typer.echo(f"✅ {message}")
                typer.echo()
                typer.echo(f"📁 워크북명: {wb['name']}")
                typer.echo(f"📍 전체경로: {wb['full_name']}")
                if saved_path:
                    typer.echo(f"💾 저장경로: {saved_path}")
                    typer.echo(f"💾 저장상태: {'저장됨' if wb['saved'] else '저장되지 않음'}")
                else:
                    typer.echo(f"⚠️  저장되지 않은 새 워크북 (필요시 직접 저장하세요)")

                typer.echo(f"📄 시트 수: {wb['sheet_count']}")
                typer.echo(f"📑 활성 시트: {wb['active_sheet']}")

                typer.echo()
                typer.echo("📋 생성된 시트:")
                for i, sheet in enumerate(wb["sheets"], 1):
                    active_mark = " (활성)" if sheet.get("is_active") else ""
                    if "error" in sheet:
                        typer.echo(f"  {i}. {sheet['name']}{active_mark} - ❌ {sheet['error']}")
                    else:
                        typer.echo(f"  {i}. {sheet['name']}{active_mark}")

                if not saved_path:
                    typer.echo()
                    typer.echo("💡 워크북을 저장하려면 Excel에서 Ctrl+S를 누르거나")
                    typer.echo("   다음 명령어를 사용하세요: oa excel workbook-save")

    except RuntimeError as e:
        error_response = create_error_response(e, "workbook-create")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)

    except Exception as e:
        error_response = create_error_response(e, "workbook-create")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ 예기치 않은 오류: {str(e)}", err=True)
            typer.echo("💡 Excel이 설치되어 있는지 확인하세요.", err=True)
        raise typer.Exit(1)

    finally:
        # 리소스 정리 - 새 앱을 생성한 경우만 종료 고려
        # 저장된 워크북이 있거나 기존 앱을 사용한 경우는 앱을 종료하지 않음
        pass


if __name__ == "__main__":
    typer.run(workbook_create)
