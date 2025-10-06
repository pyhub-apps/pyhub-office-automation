"""
Excel 데이터 구조 분석 명령어 (Issue #39)
피벗테이블용 데이터 준비 상태를 평가하고 권장사항 제공
"""

import json
import sys
from typing import Optional

import typer
import xlwings as xw

from pyhub_office_automation.version import get_version

from .utils import (
    DataTransformType,
    ExecutionTimer,
    ExpandMode,
    OutputFormat,
    analyze_data_structure,
    create_error_response,
    create_success_response,
    get_or_open_workbook,
    get_range,
    get_sheet,
    normalize_path,
    parse_range,
    validate_range_string,
)


def data_analyze(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="분석할 Excel 파일의 절대 경로"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help="열린 워크북 이름으로 접근"),
    range_str: str = typer.Option(..., "--range", help="분석할 셀 범위 (예: A1:C10, Sheet1!A1:C10)"),
    sheet: Optional[str] = typer.Option(None, "--sheet", help="시트 이름 (미지정시 활성 시트 사용)"),
    expand: Optional[ExpandMode] = typer.Option(None, "--expand", help="범위 확장 모드 (table, down, right)"),
    output_format: OutputFormat = typer.Option(OutputFormat.JSON, "--format", help="출력 형식 선택"),
    visible: bool = typer.Option(False, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부"),
):
    """
    Excel 데이터 구조를 분석하여 피벗테이블 준비 상태를 평가합니다.

    데이터가 피벗테이블에 적합한지 확인하고, 필요한 변환작업과 권장사항을 제공합니다.

    \b
    워크북 접근 방법:
      • 옵션 없음: 활성 워크북 자동 사용 (기본값)
      • --file-path: 파일 경로로 워크북 열기
      • --workbook-name: 열린 워크북 이름으로 접근

    \b
    분석 항목:
      • 교차표 형식 (Cross-tab): 월/분기가 열로 배치된 경우
      • 다단계 헤더: 중첩된 헤더 구조
      • 병합된 셀: 빈 셀로 인한 데이터 불일치
      • 소계 혼재: 데이터와 소계가 섞여있는 경우
      • 넓은 형식: 여러 지표가 열로 나열된 경우

    \b
    범위 확장 모드:
      • table: 연결된 데이터 테이블 전체로 확장
      • down: 아래쪽으로 데이터가 있는 곳까지 확장
      • right: 오른쪽으로 데이터가 있는 곳까지 확장

    \b
    사용 예제:
      oa excel data-analyze --file-path "report.xlsx" --range "A1:Z100"
      oa excel data-analyze --range "A1" --expand table
      oa excel data-analyze --workbook-name "Sales.xlsx" --range "Sheet1!A1:L100"
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

            # 범위 가져오기
            range_obj = get_range(target_sheet, parsed_range, expand)

            # 데이터 구조 분석
            analysis_result = analyze_data_structure(range_obj)

            # 추가 메타데이터
            analysis_result["source_info"] = {
                "range": range_obj.address,
                "sheet": target_sheet.name,
                "workbook": normalize_path(book.name) if hasattr(book, "name") else "Unknown",
            }

            # 변환 권장사항 추가
            recommended_transforms = []
            if "merged_cells" in analysis_result["issues"]:
                recommended_transforms.append(
                    {
                        "type": DataTransformType.UNMERGE.value,
                        "description": "병합된 셀을 해제하고 빈 값을 채워넣습니다",
                        "priority": 1,
                    }
                )

            if "subtotals_mixed" in analysis_result["issues"]:
                recommended_transforms.append(
                    {
                        "type": DataTransformType.REMOVE_SUBTOTALS.value,
                        "description": "소계 행을 제거하여 순수 데이터만 남깁니다",
                        "priority": 1,
                    }
                )

            if "multi_level_headers" in analysis_result["issues"]:
                recommended_transforms.append(
                    {
                        "type": DataTransformType.FLATTEN_HEADERS.value,
                        "description": "다단계 헤더를 단일 헤더로 결합합니다",
                        "priority": 2,
                    }
                )

            if "cross_tab" in analysis_result["issues"] or "wide_format" in analysis_result["issues"]:
                recommended_transforms.append(
                    {
                        "type": DataTransformType.UNPIVOT.value,
                        "description": "교차표나 넓은 형식을 세로 형식으로 변환합니다",
                        "priority": 3,
                    }
                )

            if len(recommended_transforms) > 1:
                recommended_transforms.append(
                    {
                        "type": DataTransformType.AUTO.value,
                        "description": "모든 필요한 변환을 자동으로 적용합니다",
                        "priority": 0,
                    }
                )

            analysis_result["recommended_transforms"] = recommended_transforms

            # 다음 단계 안내 추가
            next_steps = []
            if analysis_result["transformation_needed"]:
                next_steps.append("oa excel data-transform 명령어로 데이터를 변환하세요")
                if recommended_transforms:
                    best_transform = min(recommended_transforms, key=lambda x: x["priority"])
                    next_steps.append(f"추천: --transform-type {best_transform['type']}")

            if analysis_result["pivot_ready"]:
                next_steps.append("oa excel pivot-create 명령어로 피벗테이블을 생성할 수 있습니다")
            else:
                next_steps.append("데이터 변환 후 피벗테이블 생성이 가능합니다")

            analysis_result["next_steps"] = next_steps

            # 성공 응답 생성
            response = create_success_response(
                data=analysis_result,
                command="data-analyze",
                message=f"데이터 구조 분석이 완료되었습니다 (신뢰도: {analysis_result['confidence_score']})",
                execution_time_ms=timer.execution_time_ms,
                book=book,
                range_obj=range_obj,
            )

            # 출력 형식에 따른 결과 반환
            if output_format == OutputFormat.JSON:
                typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
            else:  # text 형식
                typer.echo(f"📊 Excel 데이터 구조 분석 결과")
                typer.echo(f"📄 파일: {analysis_result['source_info']['workbook']}")
                typer.echo(f"📋 시트: {analysis_result['source_info']['sheet']}")
                typer.echo(f"📍 범위: {analysis_result['source_info']['range']}")
                typer.echo(
                    f"📏 데이터 크기: {analysis_result['data_shape']['rows']}행 × {analysis_result['data_shape']['columns']}열"
                )
                typer.echo()

                # 분석 결과
                format_type = analysis_result["format_type"]
                format_names = {
                    "pivot_ready": "✅ 피벗테이블 준비완료",
                    "cross_tab": "📊 교차표 형식",
                    "wide_format": "📈 넓은 형식",
                    "multi_level_headers": "🔗 다단계 헤더",
                    "merged_cells": "🔀 병합된 셀",
                    "subtotals_mixed": "🧮 소계 혼재",
                    "unknown": "❓ 알 수 없음",
                }
                typer.echo(f"🏷️  데이터 형식: {format_names.get(format_type, format_type)}")
                typer.echo(f"🎯 피벗테이블 준비상태: {'✅ 준비완료' if analysis_result['pivot_ready'] else '❌ 변환 필요'}")
                typer.echo(f"🔧 변환 필요: {'아니오' if not analysis_result['transformation_needed'] else '예'}")
                typer.echo(f"📈 신뢰도: {analysis_result['confidence_score']} (0.0~1.0)")

                if analysis_result["issues"]:
                    typer.echo()
                    typer.echo("⚠️  발견된 문제점:")
                    issue_names = {
                        "merged_cells": "병합된 셀",
                        "cross_tab": "교차표 형식",
                        "multi_level_headers": "다단계 헤더",
                        "subtotals_mixed": "소계 혼재",
                        "wide_format": "넓은 형식",
                    }
                    for issue in analysis_result["issues"]:
                        typer.echo(f"  • {issue_names.get(issue, issue)}")

                if analysis_result["recommendations"]:
                    typer.echo()
                    typer.echo("💡 권장사항:")
                    for rec in analysis_result["recommendations"]:
                        typer.echo(f"  • {rec}")

                if recommended_transforms:
                    typer.echo()
                    typer.echo("🔧 추천 변환:")
                    for transform in sorted(recommended_transforms, key=lambda x: x["priority"]):
                        priority_icon = "🔥" if transform["priority"] == 0 else "⭐" if transform["priority"] == 1 else "📝"
                        typer.echo(f"  {priority_icon} {transform['type']}: {transform['description']}")

                if next_steps:
                    typer.echo()
                    typer.echo("🚀 다음 단계:")
                    for step in next_steps:
                        typer.echo(f"  • {step}")

                typer.echo(f"\n⏱️  분석 시간: {timer.execution_time_ms}ms")

    except FileNotFoundError as e:
        error_response = create_error_response(e, "data-analyze")
        if output_format == OutputFormat.JSON:
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ 파일을 찾을 수 없습니다: {file_path}", err=True)
        raise typer.Exit(1)

    except ValueError as e:
        error_response = create_error_response(e, "data-analyze")
        if output_format == OutputFormat.JSON:
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(1)

    except RuntimeError as e:
        error_response = create_error_response(e, "data-analyze")
        if output_format == OutputFormat.JSON:
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
            typer.echo(
                "💡 Excel이 설치되어 있는지 확인하고, 파일이 다른 프로그램에서 사용 중이지 않은지 확인하세요.", err=True
            )
        raise typer.Exit(1)

    except Exception as e:
        error_response = create_error_response(e, "data-analyze")
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
