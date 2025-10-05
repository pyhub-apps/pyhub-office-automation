"""
macOS Excel 엔진 (AppleScript 기반)

subprocess와 osascript를 사용하여 Excel for Mac을 제어합니다.
AppleScript를 통해 Excel의 네이티브 기능에 접근합니다.
"""

import json
import platform
import re
import subprocess
from pathlib import Path
from typing import Any, Dict, List, Optional

from .base import ChartInfo, ExcelEngineBase, RangeData, TableInfo, WorkbookInfo
from .exceptions import (
    AppleScriptError,
    ChartNotFoundError,
    EngineInitializationError,
    ExcelNotRunningError,
    RangeError,
    SheetNotFoundError,
    TableNotFoundError,
    WorkbookNotFoundError,
)


class MacOSEngine(ExcelEngineBase):
    """
    macOS AppleScript 기반 Excel 엔진

    osascript를 통해 Excel for Mac을 제어합니다.
    AppleScript의 Excel Object Model을 사용합니다.
    """

    def __init__(self):
        """macOS AppleScript 엔진 초기화"""
        if platform.system() != "Darwin":
            raise EngineInitializationError("MacOSEngine", "macOS 플랫폼에서만 사용 가능합니다")

        # Excel for Mac 설치 확인
        self._verify_excel_installed()

    def _verify_excel_installed(self):
        """Excel for Mac 설치 확인"""
        script = 'tell application "System Events" to exists application process "Microsoft Excel"'
        try:
            result = self._run_applescript(script)
            # 실행 중이 아니어도 설치되어 있으면 OK
        except:
            raise EngineInitializationError("MacOSEngine", "Excel for Mac이 설치되어 있지 않습니다")

    def _run_applescript(self, script: str, timeout: int = 30) -> str:
        """
        AppleScript 실행 및 결과 반환

        Args:
            script: 실행할 AppleScript 코드
            timeout: 타임아웃 (초)

        Returns:
            str: AppleScript 실행 결과

        Raises:
            AppleScriptError: 실행 실패 시
        """
        try:
            result = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=timeout, check=False)

            if result.returncode != 0:
                raise AppleScriptError(script, result.stderr.strip())

            return result.stdout.strip()

        except subprocess.TimeoutExpired:
            raise AppleScriptError(script, f"Timeout after {timeout} seconds")

        except Exception as e:
            raise AppleScriptError(script, str(e))

    def _parse_applescript_list(self, output: str) -> List:
        """AppleScript list 출력을 Python list로 변환"""
        if not output:
            return []

        # AppleScript list: "{item1, item2, item3}"
        output = output.strip()
        if output.startswith("{") and output.endswith("}"):
            output = output[1:-1]

        # 빈 리스트
        if not output:
            return []

        # 콤마로 분할 (중첩 괄호 고려)
        items = []
        current = ""
        depth = 0

        for char in output:
            if char == "{":
                depth += 1
            elif char == "}":
                depth -= 1
            elif char == "," and depth == 0:
                items.append(current.strip())
                current = ""
                continue
            current += char

        if current:
            items.append(current.strip())

        return items

    def _parse_applescript_record(self, output: str) -> Dict:
        """AppleScript record 출력을 Python dict로 변환"""
        # AppleScript record: "{key1:value1, key2:value2}"
        # 간단한 파싱 (복잡한 중첩 구조는 미지원)
        result = {}
        items = self._parse_applescript_list(output)

        for item in items:
            if ":" in item:
                key, value = item.split(":", 1)
                result[key.strip()] = value.strip()

        return result

    def _escape_applescript_string(self, text: str) -> str:
        """AppleScript 문자열 이스케이프"""
        # 백슬래시와 따옴표 이스케이프
        text = text.replace("\\", "\\\\")
        text = text.replace('"', '\\"')
        return text

    # ===========================================
    # 워크북 관리 (4개 명령어)
    # ===========================================

    def get_workbooks(self) -> List[WorkbookInfo]:
        """현재 열려있는 모든 워크북 목록 조회"""
        script = """
        tell application "Microsoft Excel"
            if (count of workbooks) = 0 then
                return ""
            end if

            set workbookList to {}
            repeat with wb in workbooks
                set wbName to name of wb
                set wbPath to full name of wb
                set wbSaved to saved of wb
                set sheetCount to count of sheets of wb
                set activeSheetName to name of active sheet of wb

                set wbInfo to wbName & "|" & wbPath & "|" & wbSaved & "|" & sheetCount & "|" & activeSheetName
                set end of workbookList to wbInfo
            end repeat

            set AppleScript's text item delimiters to "@@"
            set resultText to workbookList as text
            set AppleScript's text item delimiters to ""
            return resultText
        end tell
        """

        try:
            result = self._run_applescript(script)

            if not result:
                return []

            workbooks = []
            for wb_info in result.split("@@"):
                parts = wb_info.split("|")
                if len(parts) >= 5:
                    workbooks.append(
                        WorkbookInfo(
                            name=parts[0],
                            saved=parts[2].lower() == "true",
                            full_name=parts[1],
                            sheet_count=int(parts[3]),
                            active_sheet=parts[4],
                        )
                    )

            return workbooks

        except AppleScriptError as e:
            if "Microsoft Excel got an error" in str(e):
                raise ExcelNotRunningError()
            raise

    def get_workbook_info(self, workbook: Any) -> Dict[str, Any]:
        """워크북 상세 정보 조회"""
        # workbook은 워크북 이름 문자열로 가정
        workbook_name = workbook if isinstance(workbook, str) else str(workbook)

        script = f"""
        tell application "Microsoft Excel"
            tell workbook "{self._escape_applescript_string(workbook_name)}"
                set wbName to name
                set wbPath to full name
                set wbSaved to saved
                set sheetCount to count of sheets
                set activeSheetName to name of active sheet

                set sheetNames to {{}}
                repeat with sh in sheets
                    set end of sheetNames to name of sh
                end repeat

                set AppleScript's text item delimiters to ","
                set sheetNamesStr to sheetNames as text
                set AppleScript's text item delimiters to ""

                return wbName & "|" & wbPath & "|" & wbSaved & "|" & sheetCount & "|" & activeSheetName & "|" & sheetNamesStr
            end tell
        end tell
        """

        try:
            result = self._run_applescript(script)
            parts = result.split("|")

            return {
                "name": parts[0],
                "full_name": parts[1],
                "saved": parts[2].lower() == "true",
                "sheet_count": int(parts[3]),
                "active_sheet": parts[4],
                "sheets": parts[5].split(",") if len(parts) > 5 and parts[5] else [],
            }

        except AppleScriptError:
            raise WorkbookNotFoundError(workbook_name)

    def open_workbook(self, file_path: str, visible: bool = False) -> Any:
        """워크북 열기"""
        abs_path = str(Path(file_path).resolve())

        if not Path(abs_path).exists():
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {abs_path}")

        script = f"""
        tell application "Microsoft Excel"
            activate
            open POSIX file "{abs_path}"
            return name of active workbook
        end tell
        """

        try:
            workbook_name = self._run_applescript(script)
            return workbook_name  # macOS에서는 워크북 이름을 반환

        except AppleScriptError as e:
            raise RuntimeError(f"워크북 열기 실패: {str(e)}")

    def create_workbook(self, save_path: Optional[str] = None, visible: bool = False) -> Any:
        """새 워크북 생성"""
        script = """
        tell application "Microsoft Excel"
            activate
            set newWb to make new workbook
            return name of newWb
        end tell
        """

        try:
            workbook_name = self._run_applescript(script)

            # 저장 경로가 지정된 경우
            if save_path:
                abs_path = str(Path(save_path).resolve())
                save_script = f"""
                tell application "Microsoft Excel"
                    tell workbook "{workbook_name}"
                        save in POSIX file "{abs_path}"
                    end tell
                end tell
                """
                self._run_applescript(save_script)

            return workbook_name

        except AppleScriptError as e:
            raise RuntimeError(f"워크북 생성 실패: {str(e)}")

    # ===========================================
    # 시트 관리 (4개 명령어)
    # ===========================================

    def activate_sheet(self, workbook: Any, sheet_name: str):
        """시트 활성화"""
        workbook_name = workbook if isinstance(workbook, str) else str(workbook)

        script = f"""
        tell application "Microsoft Excel"
            tell workbook "{self._escape_applescript_string(workbook_name)}"
                activate object sheet "{self._escape_applescript_string(sheet_name)}"
            end tell
        end tell
        """

        try:
            self._run_applescript(script)
        except AppleScriptError:
            raise SheetNotFoundError(sheet_name)

    def add_sheet(self, workbook: Any, name: str, before: Optional[str] = None) -> str:
        """시트 추가"""
        workbook_name = workbook if isinstance(workbook, str) else str(workbook)

        if before:
            script = f"""
            tell application "Microsoft Excel"
                tell workbook "{self._escape_applescript_string(workbook_name)}"
                    set newSheet to make new sheet before sheet "{self._escape_applescript_string(before)}"
                    set name of newSheet to "{self._escape_applescript_string(name)}"
                    return name of newSheet
                end tell
            end tell
            """
        else:
            script = f"""
            tell application "Microsoft Excel"
                tell workbook "{self._escape_applescript_string(workbook_name)}"
                    set newSheet to make new sheet at end
                    set name of newSheet to "{self._escape_applescript_string(name)}"
                    return name of newSheet
                end tell
            end tell
            """

        try:
            return self._run_applescript(script)
        except AppleScriptError as e:
            if "already exists" in str(e):
                raise ValueError(f"시트 '{name}'이 이미 존재합니다")
            if before:
                raise SheetNotFoundError(before)
            raise

    def delete_sheet(self, workbook: Any, sheet_name: str):
        """시트 삭제"""
        workbook_name = workbook if isinstance(workbook, str) else str(workbook)

        script = f"""
        tell application "Microsoft Excel"
            tell workbook "{self._escape_applescript_string(workbook_name)}"
                if (count of sheets) = 1 then
                    error "마지막 시트는 삭제할 수 없습니다"
                end if
                delete sheet "{self._escape_applescript_string(sheet_name)}"
            end tell
        end tell
        """

        try:
            self._run_applescript(script)
        except AppleScriptError as e:
            if "마지막 시트" in str(e):
                raise RuntimeError("마지막 시트는 삭제할 수 없습니다")
            raise SheetNotFoundError(sheet_name)

    def rename_sheet(self, workbook: Any, old_name: str, new_name: str):
        """시트 이름 변경"""
        workbook_name = workbook if isinstance(workbook, str) else str(workbook)

        script = f"""
        tell application "Microsoft Excel"
            tell workbook "{self._escape_applescript_string(workbook_name)}"
                set name of sheet "{self._escape_applescript_string(old_name)}" to "{self._escape_applescript_string(new_name)}"
            end tell
        end tell
        """

        try:
            self._run_applescript(script)
        except AppleScriptError as e:
            if "already exists" in str(e):
                raise ValueError(f"시트 '{new_name}'이 이미 존재합니다")
            raise SheetNotFoundError(old_name)

    # ===========================================
    # 데이터 읽기/쓰기 (2개 명령어)
    # ===========================================

    def read_range(
        self, workbook: Any, sheet: str, range_str: str, expand: Optional[str] = None, include_formulas: bool = True
    ) -> RangeData:
        """셀 범위 데이터 읽기"""
        workbook_name = workbook if isinstance(workbook, str) else str(workbook)

        # 범위 확장 처리 (간단한 구현)
        if expand:
            # AppleScript에서 CurrentRegion 등을 사용
            if expand.lower() == "table":
                range_str = f"{range_str}"  # TODO: CurrentRegion 구현

        # 값 읽기
        value_script = f"""
        tell application "Microsoft Excel"
            tell workbook "{self._escape_applescript_string(workbook_name)}"
                tell sheet "{self._escape_applescript_string(sheet)}"
                    set rangeObj to range "{range_str}"
                    set cellValues to value of rangeObj
                    set cellAddress to address of rangeObj
                    set rowCnt to count of rows of rangeObj
                    set colCnt to count of columns of rangeObj

                    return cellAddress & "|" & rowCnt & "|" & colCnt & "|" & cellValues
                end tell
            end tell
        end tell
        """

        try:
            result = self._run_applescript(value_script)
            parts = result.split("|", 3)

            address = parts[0]
            row_count = int(parts[1])
            column_count = int(parts[2])
            values_str = parts[3] if len(parts) > 3 else ""

            # 값 파싱 (간단한 구현)
            values = values_str

            # 공식 읽기
            formulas = None
            if include_formulas:
                formula_script = f"""
                tell application "Microsoft Excel"
                    tell workbook "{self._escape_applescript_string(workbook_name)}"
                        tell sheet "{self._escape_applescript_string(sheet)}"
                            return formula of range "{range_str}"
                        end tell
                    end tell
                end tell
                """
                try:
                    formulas = self._run_applescript(formula_script)
                except:
                    pass

            return RangeData(
                values=values,
                formulas=formulas,
                address=address,
                sheet_name=sheet,
                row_count=row_count,
                column_count=column_count,
                cells_count=row_count * column_count,
            )

        except AppleScriptError:
            raise RangeError(range_str, "범위를 읽을 수 없습니다")

    def write_range(self, workbook: Any, sheet: str, range_str: str, data: Any, include_formulas: bool = False):
        """셀 범위에 데이터 쓰기"""
        workbook_name = workbook if isinstance(workbook, str) else str(workbook)

        # 데이터를 AppleScript 형식으로 변환
        if isinstance(data, list):
            # 2차원 배열
            if isinstance(data[0], list):
                data_str = str(data).replace("[", "{").replace("]", "}")
            else:
                # 1차원 배열
                data_str = "{" + ", ".join(str(item) for item in data) + "}"
        else:
            # 단일 값
            data_str = str(data)

        if include_formulas and isinstance(data, str) and data.startswith("="):
            # 공식 쓰기
            script = f"""
            tell application "Microsoft Excel"
                tell workbook "{self._escape_applescript_string(workbook_name)}"
                    tell sheet "{self._escape_applescript_string(sheet)}"
                        set formula of range "{range_str}" to "{self._escape_applescript_string(data)}"
                    end tell
                end tell
            end tell
            """
        else:
            # 값 쓰기
            script = f"""
            tell application "Microsoft Excel"
                tell workbook "{self._escape_applescript_string(workbook_name)}"
                    tell sheet "{self._escape_applescript_string(sheet)}"
                        set value of range "{range_str}" to {data_str}
                    end tell
                end tell
            end tell
            """

        try:
            self._run_applescript(script)
        except AppleScriptError:
            raise RangeError(range_str, "범위에 쓸 수 없습니다")

    # ===========================================
    # 테이블 (5개 명령어) - 기본 구현
    # ===========================================

    def list_tables(self, workbook: Any, sheet: Optional[str] = None) -> List[TableInfo]:
        """테이블 목록 조회 (Excel for Mac의 Table 지원)"""
        # Excel for Mac의 Table 객체 사용
        # 간단한 구현 (상세 정보는 제한적)
        return []

    def read_table(
        self,
        workbook: Any,
        table_name: str,
        columns: Optional[List[str]] = None,
        limit: Optional[int] = None,
        offset: int = 0,
    ) -> Dict[str, Any]:
        """테이블 데이터 읽기"""
        raise NotImplementedError("macOS에서 테이블 읽기는 제한적으로 지원됩니다")

    def write_table(self, workbook: Any, sheet: str, table_name: str, data: List[List[Any]], start_cell: str = "A1"):
        """테이블에 데이터 쓰기"""
        # 일반 범위로 쓰기
        self.write_range(workbook, sheet, start_cell, data)

    def analyze_table(self, workbook: Any, table_name: str) -> Dict[str, Any]:
        """테이블 데이터 분석"""
        raise NotImplementedError("macOS에서 테이블 분석은 제한적으로 지원됩니다")

    def generate_metadata(self, workbook: Any) -> Dict[str, Any]:
        """워크북 메타데이터 생성"""
        workbook_info = self.get_workbook_info(workbook)
        return {
            "workbook_name": workbook_info["name"],
            "sheet_count": workbook_info["sheet_count"],
            "sheets": [{"name": sheet} for sheet in workbook_info["sheets"]],
        }

    # ===========================================
    # 차트 (7개 명령어) - 기본 구현
    # ===========================================

    def add_chart(
        self,
        workbook: Any,
        sheet: str,
        data_range: str,
        chart_type: str,
        position: str,
        width: int = 400,
        height: int = 300,
        title: Optional[str] = None,
        **kwargs,
    ) -> str:
        """차트 생성 (Excel for Mac AppleScript 제한적 지원)"""
        workbook_name = workbook if isinstance(workbook, str) else str(workbook)

        # AppleScript 차트 타입 매핑
        chart_type_map = {
            "column": "column clustered",
            "bar": "bar clustered",
            "line": "line",
            "pie": "pie",
            "area": "area",
        }

        as_chart_type = chart_type_map.get(chart_type.lower(), "column clustered")

        script = f"""
        tell application "Microsoft Excel"
            tell workbook "{self._escape_applescript_string(workbook_name)}"
                tell sheet "{self._escape_applescript_string(sheet)}"
                    set chartObj to make new chart object with properties {{top:100, left:300, width:{width}, height:{height}}}
                    tell chart of chartObj
                        set source data to range "{data_range}"
                        set chart type to {as_chart_type}
                    end tell
                    return name of chartObj
                end tell
            end tell
        end tell
        """

        try:
            return self._run_applescript(script)
        except AppleScriptError as e:
            raise RuntimeError(f"차트 생성 실패: {str(e)}")

    def list_charts(self, workbook: Any, sheet: Optional[str] = None) -> List[ChartInfo]:
        """차트 목록 조회"""
        # Excel for Mac의 차트 객체 접근 제한적
        return []

    def configure_chart(self, workbook: Any, chart_name: str, **kwargs):
        """차트 설정"""
        workbook_name = workbook if isinstance(workbook, str) else str(workbook)

        if "title" in kwargs:
            script = f"""
            tell application "Microsoft Excel"
                tell workbook "{self._escape_applescript_string(workbook_name)}"
                    tell chart object "{self._escape_applescript_string(chart_name)}"
                        tell chart
                            set has title to true
                            set value of chart title to "{self._escape_applescript_string(kwargs["title"])}"
                        end tell
                    end tell
                end tell
            end tell
            """
            try:
                self._run_applescript(script)
            except AppleScriptError:
                raise ChartNotFoundError(chart_name)

    def position_chart(self, workbook: Any, sheet: str, chart_name: str, left: int, top: int, width: int, height: int):
        """차트 위치 및 크기 조정"""
        workbook_name = workbook if isinstance(workbook, str) else str(workbook)

        script = f"""
        tell application "Microsoft Excel"
            tell workbook "{self._escape_applescript_string(workbook_name)}"
                tell sheet "{self._escape_applescript_string(sheet)}"
                    tell chart object "{self._escape_applescript_string(chart_name)}"
                        set top to {top}
                        set left to {left}
                        set width to {width}
                        set height to {height}
                    end tell
                end tell
            end tell
        end tell
        """

        try:
            self._run_applescript(script)
        except AppleScriptError:
            raise ChartNotFoundError(chart_name)

    def export_chart(self, workbook: Any, sheet: str, chart_name: str, output_path: str, image_format: str = "png"):
        """차트 이미지 내보내기"""
        raise NotImplementedError("macOS에서 차트 내보내기는 현재 지원되지 않습니다")

    def delete_chart(self, workbook: Any, sheet: str, chart_name: str):
        """차트 삭제"""
        workbook_name = workbook if isinstance(workbook, str) else str(workbook)

        script = f"""
        tell application "Microsoft Excel"
            tell workbook "{self._escape_applescript_string(workbook_name)}"
                tell sheet "{self._escape_applescript_string(sheet)}"
                    delete chart object "{self._escape_applescript_string(chart_name)}"
                end tell
            end tell
        end tell
        """

        try:
            self._run_applescript(script)
        except AppleScriptError:
            raise ChartNotFoundError(chart_name)

    def create_pivot_chart(
        self,
        workbook: Any,
        source_sheet: str,
        source_range: str,
        dest_sheet: str,
        dest_range: str,
        chart_type: str = "column",
        **kwargs,
    ) -> str:
        """피벗 차트 생성 (macOS 제한적 지원)"""
        raise NotImplementedError("macOS에서 피벗 차트는 제한적으로 지원됩니다")
