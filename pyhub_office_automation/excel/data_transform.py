"""
Excel 데이터 변환 명령어 (Issue #39)
피벗테이블용 형식으로 데이터를 변환하는 기능 제공
"""

import json
import sys
from typing import Optional

import pandas as pd
import typer
import xlwings as xw

from pyhub_office_automation.version import get_version

from .utils import (
    DataTransformType,
    ExecutionTimer,
    ExpandMode,
    OutputFormat,
    create_error_response,
    create_success_response,
    get_or_open_workbook,
    get_range,
    get_sheet,
    normalize_path,
    parse_range,
    transform_data_auto,
    transform_data_flatten_headers,
    transform_data_remove_subtotals,
    transform_data_unmerge,
    transform_data_unpivot,
    validate_range_string,
)


def data_transform(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="변환할 Excel 파일의 절대 경로"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help="열린 워크북 이름으로 접근"),
    source_range: str = typer.Option(..., "--source-range", help="변환할 원본 데이터 범위 (예: A1:C10)"),
    source_sheet: Optional[str] = typer.Option(None, "--source-sheet", help="원본 시트 이름 (미지정시 활성 시트)"),
    expand: Optional[ExpandMode] = typer.Option(None, "--expand", help="범위 확장 모드 (table, down, right)"),
    transform_type: DataTransformType = typer.Option(..., "--transform-type", help="변환 타입"),
    output_sheet: Optional[str] = typer.Option(None, "--output-sheet", help="결과를 저장할 시트 이름 (미지정시 새 시트 생성)"),
    output_range: Optional[str] = typer.Option("A1", "--output-range", help="결과 저장 시작 위치 (기본값: A1)"),
    id_columns: Optional[str] = typer.Option(None, "--id-columns", help="Unpivot 시 고정할 열 이름들 (쉼표로 구분)"),
    preserve_original: bool = typer.Option(True, "--preserve-original/--overwrite", help="원본 데이터 보존 여부"),
    output_format: OutputFormat = typer.Option(OutputFormat.JSON, "--format", help="출력 형식 선택"),
    visible: bool = typer.Option(False, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부"),
):
    """
    Excel 데이터를 피벗테이블용 형식으로 변환합니다.

    다양한 변환 타입을 지원하여 데이터를 피벗테이블에 적합한 형태로 정리합니다.

    \b
    워크북 접근 방법:
      • 옵션 없음: 활성 워크북 자동 사용 (기본값)
      • --file-path: 파일 경로로 워크북 열기
      • --workbook-name: 열린 워크북 이름으로 접근

    \b
    변환 타입:
      • unpivot: 교차표를 세로 형식으로 변환
      • unmerge: 병합된 셀을 해제하고 값 채우기
      • flatten-headers: 다단계 헤더를 단일 헤더로 결합
      • remove-subtotals: 소계 행 제거
      • auto: 자동으로 모든 필요한 변환 적용

    \b
    범위 확장 모드:
      • table: 연결된 데이터 테이블 전체로 확장
      • down: 아래쪽으로 데이터가 있는 곳까지 확장
      • right: 오른쪽으로 데이터가 있는 곳까지 확장

    \b
    사용 예제:
      oa excel data-transform --source-range "A1:M100" --transform-type unpivot --output-sheet "PivotReady"
      oa excel data-transform --source-range "A1" --expand table --transform-type auto
      oa excel data-transform --workbook-name "Sales.xlsx" --source-range "Sheet1!A1:L100" --transform-type unmerge
    """
    book = None
    try:
        # 실행 시간 측정 시작
        with ExecutionTimer() as timer:
            # 범위 문자열 유효성 검증
            if not validate_range_string(source_range):
                raise typer.BadParameter(f"잘못된 원본 범위 형식입니다: {source_range}")

            # 워크북 연결
            book = get_or_open_workbook(file_path=file_path, workbook_name=workbook_name, visible=visible)

            # 원본 시트 및 범위 파싱
            parsed_sheet, parsed_range = parse_range(source_range)
            sheet_name = parsed_sheet or source_sheet

            # 원본 시트 가져오기
            source_sheet_obj = get_sheet(book, sheet_name)

            # 원본 범위 가져오기
            source_range_obj = get_range(source_sheet_obj, parsed_range, expand)

            # 데이터를 pandas DataFrame으로 변환
            values = source_range_obj.value
            if not values:
                raise ValueError("변환할 데이터가 비어있습니다")

            # 데이터를 2차원 리스트로 정규화
            if not isinstance(values, list):
                values = [[values]]
            elif not isinstance(values[0], list):
                values = [values]

            # DataFrame 생성 (첫 번째 행을 헤더로 사용)
            df = pd.DataFrame(
                values[1:], columns=values[0] if len(values) > 1 else [f"Column_{i+1}" for i in range(len(values[0]))]
            )
            original_shape = df.shape

            # 변환 실행
            applied_transforms = []

            if transform_type == DataTransformType.UNPIVOT:
                id_vars = None
                if id_columns:
                    id_vars = [col.strip() for col in id_columns.split(",")]
                df = transform_data_unpivot(df, id_vars=id_vars)
                applied_transforms.append("unpivot")

            elif transform_type == DataTransformType.UNMERGE:
                df = transform_data_unmerge(df)
                applied_transforms.append("unmerge")

            elif transform_type == DataTransformType.FLATTEN_HEADERS:
                df = transform_data_flatten_headers(df)
                applied_transforms.append("flatten-headers")

            elif transform_type == DataTransformType.REMOVE_SUBTOTALS:
                df = transform_data_remove_subtotals(df)
                applied_transforms.append("remove-subtotals")

            elif transform_type == DataTransformType.AUTO:
                df, applied_transforms = transform_data_auto(df)

            else:
                raise ValueError(f"지원하지 않는 변환 타입입니다: {transform_type}")

            transformed_shape = df.shape

            # 결과 시트 결정 및 생성
            if output_sheet:
                # 지정된 시트 이름 사용
                target_sheet_name = output_sheet
                try:
                    target_sheet = book.sheets[target_sheet_name]
                except:
                    # 시트가 없으면 새로 생성
                    target_sheet = book.sheets.add(target_sheet_name)
            else:
                # 자동으로 새 시트 생성
                base_name = f"Transformed_{transform_type.value}"
                counter = 1
                while True:
                    try:
                        target_sheet_name = f"{base_name}_{counter}" if counter > 1 else base_name
                        target_sheet = book.sheets.add(target_sheet_name)
                        break
                    except:
                        counter += 1
                        if counter > 100:  # 무한루프 방지
                            raise RuntimeError("새 시트 생성에 실패했습니다")

            # 결과 데이터를 Excel에 쓰기
            # 헤더와 데이터를 함께 쓰기
            result_data = [df.columns.tolist()] + df.values.tolist()

            # 출력 위치 파싱
            if output_range:
                try:
                    output_cell = target_sheet.range(output_range)
                except:
                    output_cell = target_sheet.range("A1")
            else:
                output_cell = target_sheet.range("A1")

            # 데이터 쓰기
            if result_data:
                end_row = output_cell.row + len(result_data) - 1
                end_col = output_cell.column + len(result_data[0]) - 1
                target_range = target_sheet.range((output_cell.row, output_cell.column), (end_row, end_col))
                target_range.value = result_data

            # 결과 정보 구성
            transform_result = {
                "source_info": {
                    "workbook": normalize_path(book.name) if hasattr(book, "name") else "Unknown",
                    "sheet": source_sheet_obj.name,
                    "range": source_range_obj.address,
                    "original_shape": {"rows": original_shape[0], "columns": original_shape[1]},
                },
                "transformation": {
                    "type": transform_type.value,
                    "applied_transforms": applied_transforms,
                    "id_columns": id_columns.split(",") if id_columns else None,
                },
                "output_info": {
                    "sheet": target_sheet.name,
                    "range": f"{output_cell.address}:{target_sheet.range((end_row, end_col)).address}",
                    "transformed_shape": {"rows": transformed_shape[0], "columns": transformed_shape[1]},
                },
                "statistics": {
                    "original_rows": original_shape[0],
                    "original_columns": original_shape[1],
                    "transformed_rows": transformed_shape[0],
                    "transformed_columns": transformed_shape[1],
                    "data_expansion_ratio": round(transformed_shape[0] / max(original_shape[0], 1), 2),
                    "column_reduction_ratio": round(transformed_shape[1] / max(original_shape[1], 1), 2),
                },
                "next_steps": [
                    f"변환된 데이터는 '{target_sheet.name}' 시트에 저장되었습니다",
                    "oa excel pivot-create 명령어로 피벗테이블을 생성할 수 있습니다",
                    "oa excel data-analyze 명령어로 변환 결과를 재분석할 수 있습니다",
                ],
            }

            # 성공 응답 생성
            response = create_success_response(
                data=transform_result,
                command="data-transform",
                message=f"데이터 변환이 완료되었습니다 ({', '.join(applied_transforms)})",
                execution_time_ms=timer.execution_time_ms,
                book=book,
                rows_count=transformed_shape[0],
                columns_count=transformed_shape[1],
            )

            # 출력 형식에 따른 결과 반환
            if output_format == OutputFormat.JSON:
                typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
            else:  # text 형식
                typer.echo(f"🔄 Excel 데이터 변환 완료")
                typer.echo(f"📄 파일: {transform_result['source_info']['workbook']}")
                typer.echo()

                # 원본 정보
                typer.echo("📥 원본 데이터:")
                typer.echo(f"  • 시트: {transform_result['source_info']['sheet']}")
                typer.echo(f"  • 범위: {transform_result['source_info']['range']}")
                typer.echo(
                    f"  • 크기: {transform_result['source_info']['original_shape']['rows']}행 × {transform_result['source_info']['original_shape']['columns']}열"
                )

                # 변환 정보
                typer.echo()
                typer.echo("🔧 변환 정보:")
                typer.echo(f"  • 타입: {transform_result['transformation']['type']}")
                typer.echo(f"  • 적용된 변환: {', '.join(applied_transforms)}")
                if transform_result["transformation"]["id_columns"]:
                    typer.echo(f"  • 고정 열: {', '.join(transform_result['transformation']['id_columns'])}")

                # 결과 정보
                typer.echo()
                typer.echo("📤 변환 결과:")
                typer.echo(f"  • 시트: {transform_result['output_info']['sheet']}")
                typer.echo(f"  • 범위: {transform_result['output_info']['range']}")
                typer.echo(
                    f"  • 크기: {transform_result['output_info']['transformed_shape']['rows']}행 × {transform_result['output_info']['transformed_shape']['columns']}열"
                )

                # 통계
                stats = transform_result["statistics"]
                typer.echo()
                typer.echo("📊 변환 통계:")
                typer.echo(f"  • 데이터 확장비: {stats['data_expansion_ratio']}배")
                typer.echo(f"  • 열 감소비: {stats['column_reduction_ratio']}배")

                change_rows = stats["transformed_rows"] - stats["original_rows"]
                change_cols = stats["transformed_columns"] - stats["original_columns"]
                typer.echo(f"  • 행 변화: {change_rows:+d}")
                typer.echo(f"  • 열 변화: {change_cols:+d}")

                # 다음 단계
                typer.echo()
                typer.echo("🚀 다음 단계:")
                for step in transform_result["next_steps"]:
                    typer.echo(f"  • {step}")

                typer.echo(f"\n⏱️  변환 시간: {timer.execution_time_ms}ms")

    except FileNotFoundError as e:
        error_response = create_error_response(e, "data-transform")
        if output_format == OutputFormat.JSON:
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ 파일을 찾을 수 없습니다: {file_path}", err=True)
        raise typer.Exit(1)

    except ValueError as e:
        error_response = create_error_response(e, "data-transform")
        if output_format == OutputFormat.JSON:
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)

    except RuntimeError as e:
        error_response = create_error_response(e, "data-transform")
        if output_format == OutputFormat.JSON:
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
            typer.echo(
                "💡 Excel이 설치되어 있는지 확인하고, 파일이 다른 프로그램에서 사용 중이지 않은지 확인하세요.", err=True
            )
        raise typer.Exit(1)

    except Exception as e:
        error_response = create_error_response(e, "data-transform")
        if output_format == OutputFormat.JSON:
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ 예기치 않은 오류: {str(e)}", err=True)
        raise typer.Exit(1)

    finally:
        # 워크북 정리 - 활성 워크북이나 이름으로 접근한 경우 앱 종료하지 않음
        if book is not None and not visible and file_path:
            try:
                book.app.quit()
            except:
                pass
