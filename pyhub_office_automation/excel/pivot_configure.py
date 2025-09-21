"""
피벗테이블 구성 명령어
피벗테이블의 필드 설정 및 집계 함수 구성
"""

import json
import platform
import sys
from pathlib import Path
from typing import Optional

import typer
import xlwings as xw

from pyhub_office_automation.version import get_version

from .utils import (
    create_error_response,
    create_success_response,
    format_output,
    get_or_open_workbook,
    get_sheet,
    get_workbook,
    load_data_from_file,
    normalize_path,
)


def pivot_configure(
    file_path: Optional[str] = typer.Option(None, "--file-path", help="피벗테이블이 있는 Excel 파일의 절대 경로"),
    use_active: bool = typer.Option(False, "--use-active", help="현재 활성 워크북 사용"),
    workbook_name: Optional[str] = typer.Option(None, "--workbook-name", help='열린 워크북 이름으로 접근 (예: "Sales.xlsx")'),
    pivot_name: str = typer.Option(..., "--pivot-name", help="구성할 피벗테이블 이름"),
    sheet: Optional[str] = typer.Option(None, "--sheet", help="피벗테이블이 있는 시트 이름 (지정하지 않으면 자동 검색)"),
    config_file: Optional[str] = typer.Option(None, "--config-file", help="피벗테이블 구성 JSON 파일 경로"),
    row_fields: Optional[str] = typer.Option(None, "--row-fields", help='행 필드 목록 (콤마로 구분, 예: "Region,Product")'),
    column_fields: Optional[str] = typer.Option(
        None, "--column-fields", help='열 필드 목록 (콤마로 구분, 예: "Year,Quarter")'
    ),
    value_fields: Optional[str] = typer.Option(
        None, "--value-fields", help='값 필드 설정 JSON 문자열 (예: \'[{"field":"Sales","function":"Sum"}]\')'
    ),
    filter_fields: Optional[str] = typer.Option(
        None, "--filter-fields", help='필터 필드 목록 (콤마로 구분, 예: "Category,Status")'
    ),
    clear_existing: bool = typer.Option(
        False, "--clear-existing", help="기존 필드 설정을 모두 지우고 새로 설정 (기본값: False)"
    ),
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택 (json/text)"),
    visible: bool = typer.Option(False, "--visible", help="Excel 애플리케이션을 화면에 표시할지 여부 (기본값: False)"),
    save: bool = typer.Option(True, "--save", help="구성 후 파일 저장 여부 (기본값: True)"),
):
    """
    피벗테이블의 필드 배치와 집계 함수를 구성합니다.

    Windows 전용 기능으로, COM API를 통해 피벗테이블 필드를 설정합니다.
    JSON 파일 또는 개별 옵션으로 설정을 지정할 수 있습니다.

    워크북 접근 방법:
    - --file-path: 파일 경로로 워크북 열기
    - --use-active: 현재 활성 워크북 사용
    - --workbook-name: 열린 워크북 이름으로 접근

    예제:
        oa excel pivot-configure --use-active --pivot-name "PivotTable1" --row-fields "Region,Product" --value-fields '[{"field":"Sales","function":"Sum"}]'
        oa excel pivot-configure --file-path "sales.xlsx" --pivot-name "SalesPivot" --config-file "pivot_config.json"
        oa excel pivot-configure --workbook-name "Report.xlsx" --pivot-name "PivotTable1" --column-fields "Year" --clear-existing

    JSON 설정 파일 형식:
    {
        "row_fields": ["Region", "Product"],
        "column_fields": ["Year", "Quarter"],
        "value_fields": [
            {"field": "Sales", "function": "Sum"},
            {"field": "Quantity", "function": "Average"}
        ],
        "filter_fields": ["Category", "Status"]
    }

    집계 함수 옵션: Sum, Count, Average, Max, Min, Product, CountNums, StdDev, StdDevp, Var, Varp
    """
    # 입력 값 검증
    if output_format not in ["json", "text"]:
        raise ValueError(f"잘못된 출력 형식: {output_format}. 사용 가능한 형식: json, text")

    book = None

    try:
        # Windows 전용 기능 확인
        if platform.system() != "Windows":
            raise RuntimeError("피벗테이블 구성은 Windows에서만 지원됩니다. macOS에서는 수동으로 피벗테이블을 구성해주세요.")

        # 워크북 연결
        book = get_or_open_workbook(file_path=file_path, workbook_name=workbook_name, use_active=use_active, visible=visible)

        # 구성 데이터 로드
        config_data = {}

        # JSON 파일에서 구성 로드
        if config_file:
            try:
                config_data = load_data_from_file(config_file)
                if not isinstance(config_data, dict):
                    raise ValueError("구성 파일은 JSON 객체여야 합니다")
            except Exception as e:
                raise ValueError(f"구성 파일 로드 실패: {str(e)}")

        # 개별 옵션들로 구성 데이터 업데이트
        if row_fields:
            config_data["row_fields"] = [field.strip() for field in row_fields.split(",")]

        if column_fields:
            config_data["column_fields"] = [field.strip() for field in column_fields.split(",")]

        if value_fields:
            try:
                value_config = json.loads(value_fields)
                if isinstance(value_config, list):
                    config_data["value_fields"] = value_config
                else:
                    raise ValueError("value_fields는 배열이어야 합니다")
            except json.JSONDecodeError as e:
                raise ValueError(f"value_fields JSON 파싱 실패: {str(e)}")

        if filter_fields:
            config_data["filter_fields"] = [field.strip() for field in filter_fields.split(",")]

        # 구성 데이터 검증
        if not config_data:
            raise ValueError("구성할 필드가 지정되지 않았습니다. --config-file 또는 개별 필드 옵션을 사용하세요")

        # 피벗테이블 찾기
        target_sheet = None
        pivot_table = None

        # 특정 시트가 지정된 경우
        if sheet:
            target_sheet = get_sheet(book, sheet)
            try:
                pivot_table = target_sheet.api.PivotTables(pivot_name)
            except:
                raise ValueError(f"시트 '{sheet}'에서 피벗테이블 '{pivot_name}'을 찾을 수 없습니다")
        else:
            # 전체 워크북에서 피벗테이블 검색
            for ws in book.sheets:
                try:
                    pivot_table = ws.api.PivotTables(pivot_name)
                    target_sheet = ws
                    break
                except:
                    continue

            if not pivot_table:
                raise ValueError(f"피벗테이블 '{pivot_name}'을 찾을 수 없습니다")

        # xlwings constants import
        try:
            from xlwings.constants import ConsolidationFunction, PivotFieldOrientation
        except ImportError:
            raise RuntimeError("xlwings.constants 모듈을 가져올 수 없습니다. xlwings 최신 버전이 필요합니다.")

        # 집계 함수 매핑
        function_map = {
            "Sum": ConsolidationFunction.xlSum,
            "Count": ConsolidationFunction.xlCount,
            "Average": ConsolidationFunction.xlAverage,
            "Max": ConsolidationFunction.xlMax,
            "Min": ConsolidationFunction.xlMin,
            "Product": ConsolidationFunction.xlProduct,
            "CountNums": ConsolidationFunction.xlCountNums,
            "StdDev": ConsolidationFunction.xlStdDev,
            "StdDevp": ConsolidationFunction.xlStdDevP,
            "Var": ConsolidationFunction.xlVar,
            "Varp": ConsolidationFunction.xlVarP,
        }

        configuration_results = {
            "pivot_name": pivot_name,
            "sheet": target_sheet.name,
            "configured_fields": {},
            "errors": [],
            "warnings": [],
        }

        # 기존 필드 정리 (선택적)
        if clear_existing:
            try:
                # 모든 필드를 숨김 영역으로 이동
                for field in pivot_table.PivotFields():
                    try:
                        field.Orientation = PivotFieldOrientation.xlHidden
                    except:
                        pass

                # 데이터 필드 제거
                while pivot_table.DataFields.Count > 0:
                    try:
                        pivot_table.DataFields(1).Delete()
                    except:
                        break

                configuration_results["cleared_existing"] = True
            except Exception as e:
                configuration_results["warnings"].append(f"기존 필드 정리 중 오류: {str(e)}")

        # 행 필드 설정
        if config_data.get("row_fields"):
            configured_row_fields = []
            for field_name in config_data["row_fields"]:
                try:
                    pivot_field = pivot_table.PivotFields(field_name)
                    pivot_field.Orientation = PivotFieldOrientation.xlRowField
                    configured_row_fields.append(field_name)
                except Exception as e:
                    configuration_results["errors"].append(f"행 필드 '{field_name}' 설정 실패: {str(e)}")

            configuration_results["configured_fields"]["row_fields"] = configured_row_fields

        # 열 필드 설정
        if config_data.get("column_fields"):
            configured_column_fields = []
            for field_name in config_data["column_fields"]:
                try:
                    pivot_field = pivot_table.PivotFields(field_name)
                    pivot_field.Orientation = PivotFieldOrientation.xlColumnField
                    configured_column_fields.append(field_name)
                except Exception as e:
                    configuration_results["errors"].append(f"열 필드 '{field_name}' 설정 실패: {str(e)}")

            configuration_results["configured_fields"]["column_fields"] = configured_column_fields

        # 필터 필드 설정
        if config_data.get("filter_fields"):
            configured_filter_fields = []
            for field_name in config_data["filter_fields"]:
                try:
                    pivot_field = pivot_table.PivotFields(field_name)
                    pivot_field.Orientation = PivotFieldOrientation.xlPageField
                    configured_filter_fields.append(field_name)
                except Exception as e:
                    configuration_results["errors"].append(f"필터 필드 '{field_name}' 설정 실패: {str(e)}")

            configuration_results["configured_fields"]["filter_fields"] = configured_filter_fields

        # 값 필드 설정
        if config_data.get("value_fields"):
            configured_value_fields = []
            for value_config in config_data["value_fields"]:
                if not isinstance(value_config, dict) or "field" not in value_config:
                    configuration_results["errors"].append(f"잘못된 값 필드 구성: {value_config}")
                    continue

                field_name = value_config["field"]
                function_name = value_config.get("function", "Sum")

                if function_name not in function_map:
                    configuration_results["errors"].append(f"지원되지 않는 집계 함수: {function_name}")
                    continue

                try:
                    # 값 필드 추가
                    data_field = pivot_table.AddDataField(pivot_table.PivotFields(field_name))
                    data_field.Function = function_map[function_name]

                    # 사용자 지정 이름 설정 (선택적)
                    if "name" in value_config:
                        data_field.Name = value_config["name"]

                    configured_value_fields.append({"field": field_name, "function": function_name, "name": data_field.Name})

                except Exception as e:
                    configuration_results["errors"].append(f"값 필드 '{field_name}' 설정 실패: {str(e)}")

            configuration_results["configured_fields"]["value_fields"] = configured_value_fields

        # 피벗테이블 새로고침
        try:
            pivot_table.RefreshTable()
            configuration_results["refreshed"] = True
        except Exception as e:
            configuration_results["warnings"].append(f"피벗테이블 새로고침 실패: {str(e)}")

        # 파일 저장
        save_success = False
        save_error = None
        if save:
            try:
                book.save()
                save_success = True
            except Exception as e:
                save_error = str(e)

        # 응답 데이터 구성
        data_content = {
            "configuration": configuration_results,
            "input_config": config_data,
            "success_count": sum(
                len(fields) if isinstance(fields, list) else 1
                for fields in configuration_results["configured_fields"].values()
            ),
            "error_count": len(configuration_results["errors"]),
            "warning_count": len(configuration_results["warnings"]),
            "file_info": {
                "path": (
                    str(Path(normalize_path(file_path)).resolve())
                    if file_path
                    else (normalize_path(book.fullname) if hasattr(book, "fullname") else None)
                ),
                "name": Path(normalize_path(file_path)).name if file_path else normalize_path(book.name),
                "saved": save_success,
            },
        }

        if save_error:
            data_content["save_error"] = save_error

        # 성공 메시지 구성
        message = f"피벗테이블 '{pivot_name}' 구성 완료: {data_content['success_count']}개 필드 설정됨"
        if data_content["error_count"] > 0:
            message += f" ({data_content['error_count']}개 오류)"

        response = create_success_response(data=data_content, command="pivot-configure", message=message)

        # 출력 형식에 따른 결과 반환
        if output_format == "json":
            typer.echo(json.dumps(response, ensure_ascii=False, indent=2))
        else:  # text 형식
            typer.echo(f"✅ 피벗테이블 구성 완료")
            typer.echo(f"📋 피벗테이블 이름: {pivot_name}")
            typer.echo(f"📄 파일: {data_content['file_info']['name']}")
            typer.echo(f"📍 시트: {target_sheet.name}")
            typer.echo(f"✅ 설정된 필드: {data_content['success_count']}개")

            # 구성된 필드들 표시
            config_fields = configuration_results["configured_fields"]
            if config_fields.get("row_fields"):
                typer.echo(f"   📊 행 필드: {', '.join(config_fields['row_fields'])}")
            if config_fields.get("column_fields"):
                typer.echo(f"   📊 열 필드: {', '.join(config_fields['column_fields'])}")
            if config_fields.get("filter_fields"):
                typer.echo(f"   🔍 필터 필드: {', '.join(config_fields['filter_fields'])}")
            if config_fields.get("value_fields"):
                value_info = [f"{vf['field']} ({vf['function']})" for vf in config_fields["value_fields"]]
                typer.echo(f"   📈 값 필드: {', '.join(value_info)}")

            # 오류 및 경고 표시
            if configuration_results["errors"]:
                typer.echo(f"\n❌ 오류 ({len(configuration_results['errors'])}개):")
                for error in configuration_results["errors"]:
                    typer.echo(f"   {error}")

            if configuration_results["warnings"]:
                typer.echo(f"\n⚠️ 경고 ({len(configuration_results['warnings'])}개):")
                for warning in configuration_results["warnings"]:
                    typer.echo(f"   {warning}")

            if save_success:
                typer.echo("\n💾 파일이 저장되었습니다")
            elif save:
                typer.echo(f"\n⚠️ 저장 실패: {save_error}")

            typer.echo("\n💡 피벗테이블 새로고침은 'oa excel pivot-refresh' 명령어를 사용하세요")

    except ValueError as e:
        error_response = create_error_response(e, "pivot-configure")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
        sys.exit(1)

    except RuntimeError as e:
        error_response = create_error_response(e, "pivot-configure")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ {str(e)}", err=True)
            if "Windows" in str(e):
                typer.echo(
                    "💡 피벗테이블 구성은 Windows에서만 지원됩니다. macOS에서는 Excel의 수동 기능을 사용해주세요.", err=True
                )
            else:
                typer.echo("💡 Excel이 설치되어 있는지 확인하고, xlwings 최신 버전을 사용하는지 확인하세요.", err=True)
        sys.exit(1)

    except Exception as e:
        error_response = create_error_response(e, "pivot-configure")
        if output_format == "json":
            typer.echo(json.dumps(error_response, ensure_ascii=False, indent=2), err=True)
        else:
            typer.echo(f"❌ 예기치 않은 오류: {str(e)}", err=True)
        sys.exit(1)

    finally:
        # 워크북 정리 - 활성 워크북이나 이름으로 접근한 경우 앱 종료하지 않음
        if book and not visible and file_path:
            try:
                book.app.quit()
            except:
                pass


if __name__ == "__main__":
    pivot_configure()
