"""
Excel Table 메타데이터 관리 시스템 (Issue #59)
Metadata 시트 자동 생성 및 메타데이터 CRUD 기능 제공
"""

import datetime
import platform
from pathlib import Path
from typing import Dict, List, Optional, Union

import xlwings as xw

from .utils import coords_to_excel_address

# 메타데이터 시트의 표준 구조 정의
METADATA_SHEET_NAME = "Metadata"
METADATA_TABLE_NAME = "MetadataTable"
METADATA_HEADERS = [
    "Table_Name",  # Excel Table 이름
    "Sheet_Name",  # 시트명
    "Description",  # 테이블 설명
    "Data_Type",  # 데이터 유형 카테고리
    "Column_Info",  # 주요 컬럼 정보
    "Row_Count",  # 데이터 행 수
    "Last_Updated",  # 마지막 업데이트
    "Tags",  # 태그 (쉼표 구분)
    "Notes",  # 추가 메모
]


def ensure_metadata_sheet(workbook: xw.Book) -> xw.Sheet:
    """
    Metadata 시트가 없으면 자동 생성하고 표준 구조를 설정합니다.

    Args:
        workbook: xlwings Book 객체

    Returns:
        Metadata 시트 객체

    Raises:
        Exception: 시트 생성 실패 시
    """
    try:
        # 기존 Metadata 시트 확인
        try:
            metadata_sheet = workbook.sheets[METADATA_SHEET_NAME]
            return metadata_sheet
        except:
            pass  # 시트가 없으면 새로 생성

        # 새 Metadata 시트 생성 (맨 마지막에 추가)
        metadata_sheet = workbook.sheets.add(METADATA_SHEET_NAME, after=-1)

        # 표준 헤더 설정
        header_range = metadata_sheet.range("A1").expand("right").resize(1, len(METADATA_HEADERS))
        header_range.value = METADATA_HEADERS

        # Windows에서만 Excel Table 생성
        if platform.system() == "Windows":
            try:
                # 헤더 행만으로 Excel Table 생성 (데이터 행은 나중에 추가)
                table_range = metadata_sheet.range(f"A1:{coords_to_excel_address(1, len(METADATA_HEADERS))}")

                # COM API를 통해 Excel Table 생성
                list_objects = metadata_sheet.api.ListObjects
                excel_table = list_objects.Add(
                    SourceType=1, Source=table_range.api, XlListObjectHasHeaders=1  # xlSrcRange  # xlYes
                )
                excel_table.Name = METADATA_TABLE_NAME
                excel_table.TableStyle = "TableStyleMedium2"

            except Exception as e:
                # Table 생성 실패 시 기본 헤더만 설정
                pass

        # 헤더 스타일링 (기본 포매팅)
        header_range.api.Font.Bold = True
        if platform.system() == "Windows":
            header_range.api.Interior.Color = 15773696  # 연한 파란색

        # 열 너비 자동 조정
        for col_idx in range(len(METADATA_HEADERS)):
            col_letter = coords_to_excel_address(1, col_idx + 1)[:-1]  # 행 번호 제거
            metadata_sheet.range(f"{col_letter}:{col_letter}").autofit()

        return metadata_sheet

    except Exception as e:
        raise Exception(f"Metadata 시트 생성 실패: {str(e)}")


def get_metadata_table_range(metadata_sheet: xw.Sheet) -> Optional[xw.Range]:
    """
    Metadata 시트에서 메타데이터 테이블 범위를 가져옵니다.

    Args:
        metadata_sheet: Metadata 시트 객체

    Returns:
        테이블 범위 또는 None
    """
    try:
        # Windows에서 Excel Table로 관리되는 경우
        if platform.system() == "Windows":
            try:
                for table in metadata_sheet.api.ListObjects():
                    if table.Name == METADATA_TABLE_NAME:
                        return metadata_sheet.range(table.Range.Address)
            except:
                pass

        # Table이 없는 경우 사용된 범위로 추정
        used_range = metadata_sheet.used_range
        if used_range:
            # 헤더가 올바른지 확인
            headers = used_range.rows(1).value
            if isinstance(headers, list) and len(headers) >= len(METADATA_HEADERS):
                # 헤더 일치 여부 확인
                if all(header in headers for header in METADATA_HEADERS[:3]):  # 최소 3개 헤더 확인
                    return used_range

        return None

    except Exception:
        return None


def read_metadata_records(workbook: xw.Book) -> List[Dict[str, Union[str, int, float]]]:
    """
    Metadata 시트에서 모든 메타데이터 레코드를 읽어옵니다.

    Args:
        workbook: xlwings Book 객체

    Returns:
        메타데이터 레코드 리스트
    """
    try:
        # Metadata 시트 확인
        try:
            metadata_sheet = workbook.sheets[METADATA_SHEET_NAME]
        except:
            return []  # Metadata 시트가 없으면 빈 리스트 반환

        # 테이블 범위 가져오기
        table_range = get_metadata_table_range(metadata_sheet)
        if not table_range or table_range.rows.count <= 1:
            return []  # 헤더만 있고 데이터가 없는 경우

        # 데이터 읽기
        values = table_range.value
        if not values or len(values) <= 1:
            return []

        # 헤더와 데이터 분리
        headers = values[0]
        data_rows = values[1:]

        # 딕셔너리 형태로 변환
        records = []
        for row in data_rows:
            # None 값이나 빈 행 건너뛰기
            if not row or all(cell is None or cell == "" for cell in row):
                continue

            record = {}
            for i, header in enumerate(headers):
                if i < len(row):
                    record[header] = row[i]
                else:
                    record[header] = None
            records.append(record)

        return records

    except Exception:
        return []


def write_metadata_record(
    workbook: xw.Book,
    table_name: str,
    sheet_name: str,
    description: str = "",
    data_type: str = "",
    column_info: str = "",
    row_count: int = 0,
    tags: str = "",
    notes: str = "",
) -> bool:
    """
    Metadata 시트에 새로운 메타데이터 레코드를 추가하거나 업데이트합니다.

    Args:
        workbook: xlwings Book 객체
        table_name: Excel Table 이름
        sheet_name: 시트명
        description: 테이블 설명
        data_type: 데이터 유형
        column_info: 컬럼 정보
        row_count: 행 수
        tags: 태그
        notes: 메모

    Returns:
        성공 여부
    """
    try:
        # Metadata 시트 확보 (없으면 생성)
        metadata_sheet = ensure_metadata_sheet(workbook)

        # 기존 레코드 확인
        existing_records = read_metadata_records(workbook)

        # 현재 시간
        current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # 새 레코드 데이터
        new_record = [table_name, sheet_name, description, data_type, column_info, row_count, current_time, tags, notes]

        # 기존 레코드에서 동일한 table_name 찾기
        update_row = None
        for i, record in enumerate(existing_records):
            if record.get("Table_Name") == table_name:
                update_row = i + 2  # 헤더 다음부터 시작하므로 +2
                break

        if update_row:
            # 기존 레코드 업데이트
            update_range = metadata_sheet.range(f"A{update_row}:{coords_to_excel_address(update_row, len(METADATA_HEADERS))}")
            update_range.value = new_record
        else:
            # 새 레코드 추가
            # 다음 빈 행 찾기
            table_range = get_metadata_table_range(metadata_sheet)
            if table_range:
                next_row = table_range.rows.count + 1
            else:
                next_row = 2  # 헤더 다음 행

            # 새 행에 데이터 추가
            new_range = metadata_sheet.range(f"A{next_row}:{coords_to_excel_address(next_row, len(METADATA_HEADERS))}")
            new_range.value = new_record

            # Windows에서 Excel Table 범위 확장
            if platform.system() == "Windows":
                try:
                    for table in metadata_sheet.api.ListObjects():
                        if table.Name == METADATA_TABLE_NAME:
                            # 테이블 범위를 새 행까지 확장
                            expanded_range = f"A1:{coords_to_excel_address(next_row, len(METADATA_HEADERS))}"
                            table.Resize(metadata_sheet.range(expanded_range).api)
                            break
                except:
                    pass

        return True

    except Exception as e:
        return False


def delete_metadata_record(workbook: xw.Book, table_name: str) -> bool:
    """
    Metadata 시트에서 특정 테이블의 메타데이터 레코드를 삭제합니다.

    Args:
        workbook: xlwings Book 객체
        table_name: 삭제할 Excel Table 이름

    Returns:
        성공 여부
    """
    try:
        # Metadata 시트 확인
        try:
            metadata_sheet = workbook.sheets[METADATA_SHEET_NAME]
        except:
            return False  # Metadata 시트가 없으면 실패

        # 기존 레코드 확인
        existing_records = read_metadata_records(workbook)

        # 삭제할 레코드 찾기
        delete_row = None
        for i, record in enumerate(existing_records):
            if record.get("Table_Name") == table_name:
                delete_row = i + 2  # 헤더 다음부터 시작하므로 +2
                break

        if not delete_row:
            return False  # 해당 레코드 없음

        # 행 삭제
        metadata_sheet.range(f"{delete_row}:{delete_row}").api.Delete()

        return True

    except Exception:
        return False


def get_metadata_record(workbook: xw.Book, table_name: str) -> Optional[Dict[str, Union[str, int, float]]]:
    """
    특정 테이블의 메타데이터 레코드를 조회합니다.

    Args:
        workbook: xlwings Book 객체
        table_name: Excel Table 이름

    Returns:
        메타데이터 레코드 또는 None
    """
    try:
        records = read_metadata_records(workbook)

        for record in records:
            if record.get("Table_Name") == table_name:
                return record

        return None

    except Exception:
        return None


def auto_generate_table_metadata(
    workbook: xw.Book, table_name: str, sheet_name: str
) -> Dict[str, Union[str, int, float, bool]]:
    """
    Excel Table의 메타데이터를 자동으로 분석하고 생성합니다.

    Args:
        workbook: xlwings Book 객체
        table_name: Excel Table 이름
        sheet_name: 시트명

    Returns:
        자동 생성된 메타데이터 정보
    """
    try:
        # 시트 및 테이블 가져오기
        sheet = workbook.sheets[sheet_name]

        # Excel Table 찾기
        table_range = None
        table_info = {}

        if platform.system() == "Windows":
            try:
                for table in sheet.api.ListObjects():
                    if table.Name == table_name:
                        table_range = sheet.range(table.Range.Address)
                        table_info = {
                            "range": table.Range.Address,
                            "row_count": table.Range.Rows.Count - 1,  # 헤더 제외
                            "column_count": table.Range.Columns.Count,
                            "has_headers": table.HeaderRowRange is not None,
                        }
                        break
            except:
                pass

        if not table_range:
            # Table을 찾지 못한 경우 기본값 반환
            return {
                "description": f"{table_name} 테이블",
                "data_type": "unknown",
                "column_info": "",
                "row_count": 0,
                "tags": "auto-generated",
                "notes": "자동 생성 - 테이블 분석 실패",
                "success": False,
            }

        # 데이터 분석
        values = table_range.value
        if not values or len(values) <= 1:
            return {
                "description": f"{table_name} 테이블 (데이터 없음)",
                "data_type": "empty",
                "column_info": "",
                "row_count": 0,
                "tags": "auto-generated,empty",
                "notes": "자동 생성 - 데이터 없음",
                "success": True,
            }

        # 헤더와 데이터 분리
        headers = values[0] if isinstance(values[0], list) else [values[0]]
        data_rows = values[1:] if len(values) > 1 else []

        # 컬럼 정보 생성
        column_info = ",".join([str(h) for h in headers if h is not None])

        # 데이터 타입 추론
        data_type = infer_data_type_from_columns(headers)

        # 설명 자동 생성
        description = f"{sheet_name} 시트의 {table_name} 테이블"
        if len(headers) > 0:
            description += f" ({len(headers)}개 컬럼, {len(data_rows)}행)"

        # 태그 생성
        tags = ["auto-generated"]
        if data_type != "unknown":
            tags.append(data_type)
        if len(data_rows) > 100:
            tags.append("large-dataset")

        return {
            "description": description,
            "data_type": data_type,
            "column_info": column_info[:255],  # Excel 셀 길이 제한 고려
            "row_count": len(data_rows),
            "tags": ",".join(tags),
            "notes": f"자동 생성 ({datetime.datetime.now().strftime('%Y-%m-%d %H:%M')})",
            "success": True,
        }

    except Exception as e:
        return {
            "description": f"{table_name} 테이블",
            "data_type": "unknown",
            "column_info": "",
            "row_count": 0,
            "tags": "auto-generated,error",
            "notes": f"자동 생성 실패: {str(e)}",
            "success": False,
        }


def infer_data_type_from_columns(headers: List[str]) -> str:
    """
    컬럼 이름으로부터 데이터 타입을 추론합니다.

    Args:
        headers: 컬럼 헤더 리스트

    Returns:
        추론된 데이터 타입
    """
    if not headers:
        return "unknown"

    # 컬럼명을 소문자로 변환하여 분석
    lower_headers = [str(h).lower() for h in headers if h is not None]
    header_text = " ".join(lower_headers)

    # 패턴 매칭으로 데이터 타입 추론
    if any(keyword in header_text for keyword in ["sales", "revenue", "매출", "판매", "수익"]):
        return "sales"
    elif any(keyword in header_text for keyword in ["customer", "client", "고객", "회원"]):
        return "customer"
    elif any(keyword in header_text for keyword in ["product", "item", "제품", "상품", "품목"]):
        return "product"
    elif any(keyword in header_text for keyword in ["financial", "finance", "재무", "회계", "예산"]):
        return "financial"
    elif any(keyword in header_text for keyword in ["inventory", "stock", "재고", "물류"]):
        return "inventory"
    elif any(keyword in header_text for keyword in ["employee", "hr", "staff", "직원", "인사"]):
        return "hr"
    elif any(keyword in header_text for keyword in ["order", "purchase", "주문", "구매"]):
        return "transaction"
    elif any(keyword in header_text for keyword in ["date", "time", "날짜", "시간", "기간"]):
        return "time-series"
    else:
        return "general"


def get_workbook_tables_summary(workbook: xw.Book) -> Dict[str, Union[int, List, Dict]]:
    """
    워크북의 모든 Excel Table 요약 정보와 메타데이터를 수집합니다.

    Args:
        workbook: xlwings Book 객체

    Returns:
        Tables 요약 정보와 메타데이터
    """
    try:
        summary = {"total_tables": 0, "tables_with_metadata": 0, "by_sheet": {}, "all_tables": [], "metadata_available": False}

        # 메타데이터 레코드 읽기
        metadata_records = read_metadata_records(workbook)
        metadata_dict = {record.get("Table_Name"): record for record in metadata_records if record.get("Table_Name")}
        summary["metadata_available"] = len(metadata_dict) > 0

        # 각 시트에서 Excel Table 찾기
        for sheet in workbook.sheets:
            sheet_tables = []

            try:
                if platform.system() == "Windows":
                    # Windows에서 COM API로 Table 조회
                    for table in sheet.api.ListObjects():
                        try:
                            table_info = {
                                "name": table.Name,
                                "sheet": sheet.name,
                                "range": table.Range.Address.replace("$", ""),
                                "row_count": table.Range.Rows.Count - 1,  # 헤더 제외
                                "column_count": table.Range.Columns.Count,
                            }

                            # 메타데이터 추가
                            if table.Name in metadata_dict:
                                metadata = metadata_dict[table.Name]
                                table_info["metadata"] = {
                                    "description": metadata.get("Description", ""),
                                    "data_type": metadata.get("Data_Type", ""),
                                    "tags": metadata.get("Tags", ""),
                                    "last_updated": metadata.get("Last_Updated", ""),
                                }
                                summary["tables_with_metadata"] += 1
                            else:
                                table_info["metadata"] = None

                            sheet_tables.append(table_info)
                            summary["all_tables"].append(table_info)
                            summary["total_tables"] += 1

                        except Exception:
                            continue
                else:
                    # macOS에서는 제한적인 지원
                    for table in sheet.tables:
                        try:
                            table_info = {
                                "name": table.name,
                                "sheet": sheet.name,
                                "range": table.range.address.replace("$", ""),
                                "row_count": table.range.rows.count - 1,
                                "column_count": table.range.columns.count,
                            }

                            # 메타데이터 추가
                            if table.name in metadata_dict:
                                metadata = metadata_dict[table.name]
                                table_info["metadata"] = {
                                    "description": metadata.get("Description", ""),
                                    "data_type": metadata.get("Data_Type", ""),
                                    "tags": metadata.get("Tags", ""),
                                    "last_updated": metadata.get("Last_Updated", ""),
                                }
                                summary["tables_with_metadata"] += 1
                            else:
                                table_info["metadata"] = None

                            sheet_tables.append(table_info)
                            summary["all_tables"].append(table_info)
                            summary["total_tables"] += 1

                        except Exception:
                            continue

                # 시트별 요약 추가
                if sheet_tables:
                    summary["by_sheet"][sheet.name] = {"count": len(sheet_tables), "tables": sheet_tables}

            except Exception:
                continue

        return summary

    except Exception:
        return {
            "total_tables": 0,
            "tables_with_metadata": 0,
            "by_sheet": {},
            "all_tables": [],
            "metadata_available": False,
            "error": "Tables 요약 생성 실패",
        }
