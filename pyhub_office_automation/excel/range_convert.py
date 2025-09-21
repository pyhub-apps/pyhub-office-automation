"""
Excel 셀 범위 데이터 형식 변환 명령어 (Typer 버전)
문자열에서 숫자로 변환하는 기능 제공
"""

import json
import re
from pathlib import Path
from typing import Optional

import typer
import xlwings as xw

from pyhub_office_automation.version import get_version

from .utils import (
    ExecutionTimer,
    ExpandMode,
    create_error_response,
    create_success_response,
    get_or_open_workbook,
    get_range,
    get_sheet,
    normalize_path,
    parse_range,
    validate_range_string,
)


class DataConverter:
    """데이터 형식 변환 클래스"""

    @staticmethod
    def convert_to_number(value, remove_comma=True, remove_currency=True, parse_percent=True):
        """문자열을 숫자로 변환"""
        if value is None or value == "":
            return value

        # 이미 숫자인 경우 그대로 반환
        if isinstance(value, (int, float)):
            return value

        # 문자열이 아닌 경우 문자열로 변환
        str_value = str(value).strip()

        if not str_value:
            return value

        # 원본 값 저장
        original_value = str_value

        # 쉼표 제거
        if remove_comma:
            str_value = str_value.replace(",", "")

        # 통화 기호 제거 (원, 달러, 유로 등)
        if remove_currency:
            currency_symbols = ["₩", "$", "€", "¥", "£", "원", "달러", "유로", "엔", "파운드"]
            for symbol in currency_symbols:
                str_value = str_value.replace(symbol, "")

        # 백분율 처리
        if parse_percent and str_value.endswith("%"):
            try:
                number_part = str_value[:-1].strip()
                if number_part:
                    return float(number_part) / 100
            except ValueError:
                pass

        # 괄호로 둘러싸인 음수 처리 (예: (100) -> -100)
        bracket_match = re.match(r"^\(([0-9,.]+)\)$", str_value)
        if bracket_match:
            str_value = "-" + bracket_match.group(1)

        # 공백 제거
        str_value = str_value.strip()

        # 숫자 변환 시도
        try:
            # 정수 변환 시도
            if "." not in str_value:
                return int(str_value)
            else:
                return float(str_value)
        except ValueError:
            # 변환 실패시 원본 값 반환
            return original_value


def range_convert(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="변환할 Excel 파일의 절대 경로"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help="열린 워크북 이름으로 접근"),
    range_str: str = typer.Option(
        ..., "--range", help="변환할 셀 범위 (예: A1:C10, Sheet1!A1:C10) ※단일 셀 + expand 시 오류 가능"
    ),
    sheet: Optional[str] = typer.Option(None, "--sheet", help="시트 이름 (미지정시 활성 시트 사용)"),
    expand: Optional[ExpandMode] = typer.Option(None, "--expand", help="범위 확장 모드 (table, down, right)"),
    conversion_type: str = typer.Option("auto", "--type", help="변환 유형 (auto, number, currency, percent)"),
    remove_comma: bool = typer.Option(True, "--remove-comma/--keep-comma", help="쉼표 제거 여부"),
    remove_currency: bool = typer.Option(True, "--remove-currency/--keep-currency", help="통화 기호 제거 여부"),
    parse_percent: bool = typer.Option(True, "--parse-percent/--keep-percent", help="백분율 파싱 여부"),
    save: bool = typer.Option(True, "--save/--no-save", help="변환 후 파일 저장 여부"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택"),
    visible: bool = typer.Option(False, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부"),
):
    """
    Excel 셀 범위의 문자열 데이터를 숫자로 변환합니다.

    쉼표, 통화 기호, 백분율 등이 포함된 문자열을 숫자로 변환할 수 있습니다.

    \b
    워크북 접근 방법:
      • 옵션 없음: 활성 워크북 자동 사용 (기본값)
      • --file-path: 파일 경로로 워크북 열기
      • --workbook-name: 열린 워크북 이름으로 접근

    \b
    범위 확장 모드:
      • table: 연결된 데이터 테이블 전체로 확장
      • down: 아래쪽으로 데이터가 있는 곳까지 확장
      • right: 오른쪽으로 데이터가 있는 곳까지 확장

    \b
    변환 예제:
      • "1,234" → 1234 (쉼표 제거)
      • "₩1,000" → 1000 (통화 기호 및 쉼표 제거)
      • "50%" → 0.5 (백분율을 소수로 변환)
      • "(100)" → -100 (괄호형 음수)

    \b
    주의사항:
      단일 셀(예: H2)과 --expand 옵션을 함께 사용할 때 xlwings 버그로 인해
      오류가 발생할 수 있습니다. 이런 경우 다중 셀 범위(G2:H2)를 사용하거나
      expand 없이 정확한 범위를 지정하세요.

    \b
    사용 예제:
      oa excel range-convert --range "A1:C10" --remove-comma
      oa excel range-convert --file-path "data.xlsx" --range "A1:C10" --remove-currency
      oa excel range-convert --range "G2:H2" --expand table --parse-percent
    """
    book = None
    try:
        # 실행 시간 측정 시작
        with ExecutionTimer() as timer:
            # 범위 문자열 유효성 검증
            if not validate_range_string(range_str):
                raise typer.BadParameter(f"잘못된 범위 형식입니다: {range_str}")

            # 워크북 연결
            book = get_or_open_workbook(file_path=file_path, workbook_name=workbook_name, visible=visible)

            # 시트 및 범위 파싱
            parsed_sheet, parsed_range = parse_range(range_str)
            sheet_name = parsed_sheet or sheet

            # 시트 가져오기
            target_sheet = get_sheet(book, sheet_name)

            # 범위 가져오기 (expand 적용)
            range_obj = get_range(target_sheet, parsed_range, expand)

            # 데이터 읽기
            original_values = range_obj.value

            # 변환기 생성
            converter = DataConverter()

            # 데이터 변환
            if isinstance(original_values, list):
                if len(original_values) > 0 and isinstance(original_values[0], list):
                    # 2차원 데이터
                    converted_values = []
                    for row in original_values:
                        converted_row = []
                        for cell in row:
                            converted_cell = converter.convert_to_number(cell, remove_comma, remove_currency, parse_percent)
                            converted_row.append(converted_cell)
                        converted_values.append(converted_row)
                else:
                    # 1차원 데이터
                    converted_values = []
                    for cell in original_values:
                        converted_cell = converter.convert_to_number(cell, remove_comma, remove_currency, parse_percent)
                        converted_values.append(converted_cell)
            else:
                # 단일 값
                converted_values = converter.convert_to_number(original_values, remove_comma, remove_currency, parse_percent)

            # 변환된 데이터를 다시 Excel에 쓰기
            range_obj.value = converted_values

            # 변환 통계 계산
            def count_conversions(original, converted):
                """변환된 항목 수 계산"""
                if isinstance(original, list):
                    if len(original) > 0 and isinstance(original[0], list):
                        # 2차원
                        count = 0
                        for i, row in enumerate(original):
                            for j, cell in enumerate(row):
                                if str(cell) != str(converted[i][j]):
                                    count += 1
                        return count
                    else:
                        # 1차원
                        count = 0
                        for i, cell in enumerate(original):
                            if str(cell) != str(converted[i]):
                                count += 1
                        return count
                else:
                    # 단일 값
                    return 1 if str(original) != str(converted) else 0

            conversions_count = count_conversions(original_values, converted_values)

            # 저장 처리
            saved = False
            if save:
                try:
                    book.save()
                    saved = True
                except Exception as e:
                    # 저장 실패해도 변환은 완료된 상태
                    pass

            # 변환 정보 수집
            conversion_info = {
                "range": range_obj.address,
                "sheet": target_sheet.name,
                "conversions_applied": conversions_count,
                "total_cells": range_obj.count,
                "conversion_rate": f"{(conversions_count / range_obj.count * 100):.1f}%",
                "options": {
                    "remove_comma": remove_comma,
                    "remove_currency": remove_currency,
                    "parse_percent": parse_percent,
                    "conversion_type": conversion_type,
                },
                "saved": saved,
            }

            # 워크북 정보 추가
            workbook_info = {
                "name": normalize_path(book.name),
                "full_name": normalize_path(book.fullname),
                "saved": getattr(book, "saved", True),
            }

            # 데이터 구성
            data_content = {
                "conversion": conversion_info,
                "workbook": workbook_info,
                "expand_mode": expand.value if expand else None,
            }

            # 성공 메시지 생성
            save_status = "저장됨" if saved else ("저장하지 않음" if not save else "저장 실패")
            message = f"범위 '{range_obj.address}'에서 {conversions_count}개 항목을 변환했습니다 ({save_status})"

            # 성공 응답 생성
            response = create_success_response(
                data=data_content,
                command="range-convert",
                message=message,
                execution_time_ms=timer.execution_time_ms,
                book=book,
                range_obj=range_obj,
                data_size=len(str(converted_values).encode("utf-8")),
            )

            # 출력 형식에 따른 결과 반환
            if output_format == "json":
                typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
            else:  # text 형식
                conv = conversion_info
                wb = workbook_info

                typer.echo(f"✅ {message}")
                typer.echo()
                typer.echo(f"📁 워크북: {wb['name']}")
                typer.echo(f"📄 시트: {conv['sheet']}")
                typer.echo(f"📍 범위: {conv['range']}")
                typer.echo(f"🔄 변환: {conv['conversions_applied']}/{conv['total_cells']} ({conv['conversion_rate']})")

                if saved:
                    typer.echo(f"💾 저장: ✅ 완료")
                elif not save:
                    typer.echo(f"💾 저장: ⚠️ 저장하지 않음 (--no-save 옵션)")
                else:
                    typer.echo(f"💾 저장: ❌ 실패")

    except FileNotFoundError as e:
        error_response = create_error_response(e, "range-convert")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ 파일을 찾을 수 없습니다", err=True)
        raise typer.Exit(1)

    except ValueError as e:
        error_response = create_error_response(e, "range-convert")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)

    except Exception as e:
        error_response = create_error_response(e, "range-convert")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ 예기치 않은 오류: {str(e)}", err=True)
            typer.echo(
                "💡 Excel이 설치되어 있는지 확인하고, 파일이 다른 프로그램에서 사용 중이지 않은지 확인하세요.", err=True
            )
        raise typer.Exit(1)

    finally:
        # 워크북 정리 - 활성 워크북이나 이름으로 접근한 경우 앱 종료하지 않음
        if book and not visible and file_path:
            try:
                book.app.quit()
            except:
                pass


if __name__ == "__main__":
    typer.run(range_convert)
