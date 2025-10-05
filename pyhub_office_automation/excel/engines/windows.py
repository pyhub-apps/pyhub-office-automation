"""
Windows Excel 엔진 (pywin32 COM 기반)

pywin32를 사용하여 Windows COM API로 Excel을 제어합니다.
VBA와 동등한 수준의 기능을 제공합니다.
"""

import datetime
import gc
import os
import platform
from pathlib import Path
from typing import Any, Dict, List, Optional

import pythoncom

from .base import ChartInfo, ExcelEngineBase, RangeData, TableInfo, WorkbookInfo
from .exceptions import (
    ChartNotFoundError,
    COMError,
    EngineInitializationError,
    ExcelNotRunningError,
    RangeError,
    SheetNotFoundError,
    TableNotFoundError,
    WorkbookNotFoundError,
)


class WindowsEngine(ExcelEngineBase):
    """
    Windows COM API 기반 Excel 엔진 (pywin32)

    Excel.Application COM 객체를 사용하여 Excel을 제어합니다.
    VBA의 모든 기능을 Python에서 사용할 수 있습니다.
    """

    def __init__(self):
        """Windows COM 엔진 초기화"""
        if platform.system() != "Windows":
            raise EngineInitializationError("WindowsEngine", "Windows 플랫폼에서만 사용 가능합니다")

        try:
            import win32com.client

            # COM 초기화
            pythoncom.CoInitialize()

            # Excel Application 연결 (Early Binding)
            self.xl = win32com.client.gencache.EnsureDispatch("Excel.Application")
            self._win32com = win32com.client
            self._constants = win32com.client.constants

        except ImportError:
            raise EngineInitializationError("WindowsEngine", "pywin32 패키지가 설치되지 않았습니다")

        except Exception as e:
            raise EngineInitializationError("WindowsEngine", f"COM 초기화 실패: {str(e)}")

    def __del__(self):
        """COM 정리"""
        try:
            pythoncom.CoUninitialize()
        except:
            pass

    # ===========================================
    # 워크북 관리 (4개 명령어)
    # ===========================================

    def get_workbooks(self) -> List[WorkbookInfo]:
        """현재 열려있는 모든 워크북 목록 조회"""
        try:
            workbooks = []

            if self.xl.Workbooks.Count == 0:
                return workbooks

            for wb in self.xl.Workbooks:
                try:
                    # 워크북 정보 수집
                    wb_info = WorkbookInfo(
                        name=wb.Name,
                        saved=wb.Saved,
                        full_name=wb.FullName,
                        sheet_count=wb.Sheets.Count,
                        active_sheet=wb.ActiveSheet.Name if wb.Sheets.Count > 0 else None,
                    )

                    # 파일 정보 추가
                    if os.path.exists(wb.FullName):
                        file_stat = os.stat(wb.FullName)
                        wb_info.file_size_bytes = file_stat.st_size
                        wb_info.last_modified = datetime.datetime.fromtimestamp(file_stat.st_mtime).isoformat()

                    workbooks.append(wb_info)

                except Exception as e:
                    # 개별 워크북 오류는 건너뛰기
                    continue

            return workbooks

        except Exception as e:
            raise COMError(f"워크북 목록 조회 실패: {str(e)}")

    def get_workbook_info(self, workbook: Any) -> Dict[str, Any]:
        """워크북 상세 정보 조회"""
        try:
            info = {
                "name": workbook.Name,
                "full_name": workbook.FullName,
                "saved": workbook.Saved,
                "sheet_count": workbook.Sheets.Count,
                "active_sheet": workbook.ActiveSheet.Name,
                "sheets": [sheet.Name for sheet in workbook.Sheets],
            }

            # 파일 정보
            if os.path.exists(workbook.FullName):
                file_stat = os.stat(workbook.FullName)
                info["file_size_bytes"] = file_stat.st_size
                info["last_modified"] = datetime.datetime.fromtimestamp(file_stat.st_mtime).isoformat()

            return info

        except Exception as e:
            raise COMError(f"워크북 정보 조회 실패: {str(e)}")

    def open_workbook(self, file_path: str, visible: bool = False) -> Any:
        """워크북 열기"""
        try:
            # 파일 경로 정규화
            abs_path = str(Path(file_path).resolve())

            if not os.path.exists(abs_path):
                raise FileNotFoundError(f"파일을 찾을 수 없습니다: {abs_path}")

            # Excel 표시 설정
            self.xl.Visible = visible

            # 워크북 열기
            workbook = self.xl.Workbooks.Open(abs_path)

            return workbook

        except FileNotFoundError:
            raise

        except Exception as e:
            raise COMError(f"워크북 열기 실패: {str(e)}")

    def create_workbook(self, save_path: Optional[str] = None, visible: bool = False) -> Any:
        """새 워크북 생성"""
        try:
            # Excel 표시 설정
            self.xl.Visible = visible

            # 새 워크북 생성
            workbook = self.xl.Workbooks.Add()

            # 저장 경로가 지정된 경우 저장
            if save_path:
                abs_path = str(Path(save_path).resolve())
                workbook.SaveAs(abs_path)

            return workbook

        except Exception as e:
            raise COMError(f"워크북 생성 실패: {str(e)}")

    # ===========================================
    # 시트 관리 (4개 명령어)
    # ===========================================

    def activate_sheet(self, workbook: Any, sheet_name: str):
        """시트 활성화"""
        try:
            sheet = workbook.Sheets(sheet_name)
            sheet.Activate()

        except Exception as e:
            if "Subscript out of range" in str(e):
                raise SheetNotFoundError(sheet_name)
            raise COMError(f"시트 활성화 실패: {str(e)}")

    def add_sheet(self, workbook: Any, name: str, before: Optional[str] = None) -> str:
        """시트 추가"""
        try:
            # 이름 중복 확인
            try:
                workbook.Sheets(name)
                raise ValueError(f"시트 '{name}'이 이미 존재합니다")
            except:
                pass  # 존재하지 않으면 정상

            # 시트 추가
            if before:
                before_sheet = workbook.Sheets(before)
                new_sheet = workbook.Sheets.Add(Before=before_sheet)
            else:
                # 마지막에 추가
                new_sheet = workbook.Sheets.Add(After=workbook.Sheets(workbook.Sheets.Count))

            new_sheet.Name = name

            return new_sheet.Name

        except ValueError:
            raise

        except Exception as e:
            if "Subscript out of range" in str(e):
                raise SheetNotFoundError(before)
            raise COMError(f"시트 추가 실패: {str(e)}")

    def delete_sheet(self, workbook: Any, sheet_name: str):
        """시트 삭제"""
        try:
            if workbook.Sheets.Count <= 1:
                raise RuntimeError("마지막 시트는 삭제할 수 없습니다")

            sheet = workbook.Sheets(sheet_name)

            # 경고 메시지 비활성화
            self.xl.DisplayAlerts = False
            sheet.Delete()
            self.xl.DisplayAlerts = True

        except RuntimeError:
            raise

        except Exception as e:
            self.xl.DisplayAlerts = True  # 복원
            if "Subscript out of range" in str(e):
                raise SheetNotFoundError(sheet_name)
            raise COMError(f"시트 삭제 실패: {str(e)}")

    def rename_sheet(self, workbook: Any, old_name: str, new_name: str):
        """시트 이름 변경"""
        try:
            # 새 이름 중복 확인
            try:
                workbook.Sheets(new_name)
                raise ValueError(f"시트 '{new_name}'이 이미 존재합니다")
            except:
                pass  # 존재하지 않으면 정상

            # 이름 변경
            sheet = workbook.Sheets(old_name)
            sheet.Name = new_name

        except ValueError:
            raise

        except Exception as e:
            if "Subscript out of range" in str(e):
                raise SheetNotFoundError(old_name)
            raise COMError(f"시트 이름 변경 실패: {str(e)}")

    # ===========================================
    # 데이터 읽기/쓰기 (2개 명령어)
    # ===========================================

    def read_range(
        self, workbook: Any, sheet: str, range_str: str, expand: Optional[str] = None, include_formulas: bool = True
    ) -> RangeData:
        """셀 범위 데이터 읽기"""
        try:
            ws = workbook.Sheets(sheet)
            range_obj = ws.Range(range_str)

            # 범위 확장
            if expand:
                if expand.lower() == "table":
                    range_obj = range_obj.CurrentRegion
                elif expand.lower() == "down":
                    range_obj = ws.Range(range_obj, range_obj.End(self._constants.xlDown))
                elif expand.lower() == "right":
                    range_obj = ws.Range(range_obj, range_obj.End(self._constants.xlToRight))

            # 값 읽기
            values = range_obj.Value

            # 공식 읽기
            formulas = None
            if include_formulas:
                try:
                    formulas = range_obj.Formula
                except:
                    pass

            # 범위 정보
            row_count = range_obj.Rows.Count
            column_count = range_obj.Columns.Count
            cells_count = range_obj.Count

            return RangeData(
                values=values,
                formulas=formulas,
                address=range_obj.Address,
                sheet_name=sheet,
                row_count=row_count,
                column_count=column_count,
                cells_count=cells_count,
            )

        except Exception as e:
            if "Subscript out of range" in str(e):
                raise SheetNotFoundError(sheet)
            raise RangeError(range_str, str(e))

    def write_range(self, workbook: Any, sheet: str, range_str: str, data: Any, include_formulas: bool = False):
        """셀 범위에 데이터 쓰기"""
        try:
            ws = workbook.Sheets(sheet)
            range_obj = ws.Range(range_str)

            if include_formulas and isinstance(data, str) and data.startswith("="):
                # 공식 쓰기
                range_obj.Formula = data
            else:
                # 값 쓰기
                range_obj.Value = data

        except Exception as e:
            if "Subscript out of range" in str(e):
                raise SheetNotFoundError(sheet)
            raise RangeError(range_str, str(e))

    # ===========================================
    # 테이블 (5개 명령어)
    # ===========================================

    def list_tables(self, workbook: Any, sheet: Optional[str] = None) -> List[TableInfo]:
        """테이블 목록 조회"""
        try:
            tables = []

            # 시트 범위 결정
            if sheet:
                sheets = [workbook.Sheets(sheet)]
            else:
                sheets = list(workbook.Sheets)

            for ws in sheets:
                try:
                    for tbl in ws.ListObjects:
                        # 헤더 추출
                        headers = []
                        if tbl.HeaderRowRange:
                            header_values = tbl.HeaderRowRange.Value
                            if isinstance(header_values, tuple):
                                headers = list(header_values[0]) if header_values else []
                            else:
                                headers = [header_values] if header_values else []

                        # 샘플 데이터 (최대 5행)
                        sample_data = None
                        if tbl.DataBodyRange:
                            data_range = tbl.DataBodyRange
                            sample_rows = min(5, data_range.Rows.Count)
                            if sample_rows > 0:
                                sample_range = data_range.Resize(sample_rows, data_range.Columns.Count)
                                sample_values = sample_range.Value
                                if isinstance(sample_values, tuple):
                                    sample_data = [list(row) for row in sample_values]
                                else:
                                    sample_data = [[sample_values]]

                        tables.append(
                            TableInfo(
                                name=tbl.Name,
                                sheet_name=ws.Name,
                                address=tbl.Range.Address,
                                row_count=tbl.DataBodyRange.Rows.Count if tbl.DataBodyRange else 0,
                                column_count=tbl.ListColumns.Count,
                                headers=headers,
                                sample_data=sample_data,
                            )
                        )
                except:
                    continue

            return tables

        except Exception as e:
            if sheet and "Subscript out of range" in str(e):
                raise SheetNotFoundError(sheet)
            raise COMError(f"테이블 목록 조회 실패: {str(e)}")

    def read_table(
        self,
        workbook: Any,
        table_name: str,
        columns: Optional[List[str]] = None,
        limit: Optional[int] = None,
        offset: int = 0,
    ) -> Dict[str, Any]:
        """테이블 데이터 읽기"""
        try:
            # 테이블 찾기
            table = None
            for ws in workbook.Sheets:
                try:
                    table = ws.ListObjects(table_name)
                    break
                except:
                    continue

            if not table:
                raise TableNotFoundError(table_name)

            # 헤더 추출
            headers = []
            if table.HeaderRowRange:
                header_values = table.HeaderRowRange.Value
                if isinstance(header_values, tuple):
                    headers = list(header_values[0])
                else:
                    headers = [header_values]

            # 데이터 추출
            data = []
            if table.DataBodyRange:
                data_values = table.DataBodyRange.Value

                if isinstance(data_values, tuple):
                    # 2차원 데이터
                    data = [list(row) for row in data_values]
                else:
                    # 단일 행
                    data = [[data_values]]

                # 오프셋 적용
                if offset > 0:
                    data = data[offset:]

                # 제한 적용
                if limit:
                    data = data[:limit]

            # 컬럼 필터링
            if columns:
                col_indices = [headers.index(col) for col in columns if col in headers]
                data = [[row[i] for i in col_indices] for row in data]
                headers = [headers[i] for i in col_indices]

            return {"table_name": table_name, "headers": headers, "data": data, "row_count": len(data)}

        except TableNotFoundError:
            raise

        except Exception as e:
            raise COMError(f"테이블 읽기 실패: {str(e)}")

    def write_table(self, workbook: Any, sheet: str, table_name: str, data: List[List[Any]], start_cell: str = "A1"):
        """테이블에 데이터 쓰기"""
        try:
            ws = workbook.Sheets(sheet)

            # 데이터 범위 계산
            rows = len(data)
            cols = len(data[0]) if data else 0

            # 데이터 쓰기
            if rows > 0 and cols > 0:
                start_range = ws.Range(start_cell)
                end_range = start_range.Offset(rows - 1, cols - 1)
                data_range = ws.Range(start_range, end_range)
                data_range.Value = data

                # 테이블로 변환
                try:
                    # 기존 테이블 삭제
                    try:
                        old_table = ws.ListObjects(table_name)
                        old_table.Delete()
                    except:
                        pass

                    # 새 테이블 생성
                    table = ws.ListObjects.Add(
                        SourceType=self._constants.xlSrcRange,
                        Source=data_range,
                        XlListObjectHasHeaders=self._constants.xlYes,
                    )
                    table.Name = table_name

                except Exception as e:
                    # 테이블 생성 실패는 경고만
                    pass

        except Exception as e:
            if "Subscript out of range" in str(e):
                raise SheetNotFoundError(sheet)
            raise COMError(f"테이블 쓰기 실패: {str(e)}")

    def analyze_table(self, workbook: Any, table_name: str) -> Dict[str, Any]:
        """테이블 데이터 분석"""
        # 기본 통계 정보만 제공 (상세 분석은 pandas 사용 권장)
        table_data = self.read_table(workbook, table_name)

        analysis = {
            "table_name": table_name,
            "row_count": table_data["row_count"],
            "column_count": len(table_data["headers"]),
            "headers": table_data["headers"],
            "sample_data": table_data["data"][:5] if table_data["data"] else [],
        }

        return analysis

    def generate_metadata(self, workbook: Any) -> Dict[str, Any]:
        """워크북 메타데이터 생성"""
        try:
            metadata = {
                "workbook_name": workbook.Name,
                "sheet_count": workbook.Sheets.Count,
                "sheets": [],
            }

            for sheet in workbook.Sheets:
                sheet_info = {
                    "name": sheet.Name,
                    "tables": [],
                    "used_range": sheet.UsedRange.Address if sheet.UsedRange else None,
                }

                # 테이블 정보
                try:
                    for table in sheet.ListObjects:
                        sheet_info["tables"].append({"name": table.Name, "address": table.Range.Address})
                except:
                    pass

                metadata["sheets"].append(sheet_info)

            return metadata

        except Exception as e:
            raise COMError(f"메타데이터 생성 실패: {str(e)}")

    # ===========================================
    # 차트 (7개 명령어)
    # ===========================================

    # 차트 타입 매핑 (Excel COM 상수)
    CHART_TYPE_MAP = {
        "column": 51,  # xlColumnClustered
        "column_clustered": 51,
        "column_stacked": 52,  # xlColumnStacked
        "column_stacked_100": 53,  # xlColumnStacked100
        "bar": 57,  # xlBarClustered
        "bar_clustered": 57,
        "bar_stacked": 58,  # xlBarStacked
        "bar_stacked_100": 59,  # xlBarStacked100
        "line": 4,  # xlLine
        "line_markers": 65,  # xlLineMarkers
        "pie": 5,  # xlPie
        "doughnut": -4120,  # xlDoughnut
        "area": 1,  # xlArea
        "area_stacked": 76,  # xlAreaStacked
        "area_stacked_100": 77,  # xlAreaStacked100
        "scatter": -4169,  # xlXYScatter
        "scatter_lines": 74,  # xlXYScatterLines
        "scatter_smooth": 72,  # xlXYScatterSmooth
        "bubble": 15,  # xlBubble
        "combo": -4111,  # xlCombination
        "map": 140,  # xlRegionMap (Map Chart)
    }

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
        """차트 생성"""
        try:
            ws = workbook.Sheets(sheet)

            # 차트 타입 상수 가져오기
            chart_type_code = self.CHART_TYPE_MAP.get(chart_type.lower(), 51)

            # 위치 셀의 좌표 가져오기
            position_cell = ws.Range(position)
            left = position_cell.Left
            top = position_cell.Top

            # 차트 객체 생성
            chart_obj = ws.ChartObjects().Add(Left=left, Top=top, Width=width, Height=height)

            # 차트 설정
            chart = chart_obj.Chart
            chart.SetSourceData(ws.Range(data_range))
            chart.ChartType = chart_type_code

            # 제목 설정
            if title:
                chart.HasTitle = True
                chart.ChartTitle.Text = title

            # 범례 설정
            legend_position = kwargs.get("legend_position")
            if legend_position:
                legend_map = {
                    "top": -4160,  # xlTop
                    "bottom": -4107,  # xlBottom
                    "left": -4131,  # xlLeft
                    "right": -4152,  # xlRight
                }
                if legend_position.lower() in legend_map:
                    chart.HasLegend = True
                    chart.Legend.Position = legend_map[legend_position.lower()]
                elif legend_position.lower() == "none":
                    chart.HasLegend = False

            # 데이터 레이블
            if kwargs.get("show_data_labels"):
                try:
                    for series in chart.SeriesCollection():
                        series.HasDataLabels = True
                except:
                    pass

            return chart_obj.Name

        except Exception as e:
            if "Subscript out of range" in str(e):
                raise SheetNotFoundError(sheet)
            raise COMError(f"차트 생성 실패: {str(e)}")

    def list_charts(self, workbook: Any, sheet: Optional[str] = None) -> List[ChartInfo]:
        """차트 목록 조회"""
        try:
            charts = []

            # 시트 범위 결정
            if sheet:
                sheets = [workbook.Sheets(sheet)]
            else:
                sheets = list(workbook.Sheets)

            for ws in sheets:
                try:
                    for chart_obj in ws.ChartObjects():
                        chart = chart_obj.Chart

                        # 차트 타입 역매핑
                        chart_type_name = "unknown"
                        for name, code in self.CHART_TYPE_MAP.items():
                            if chart.ChartType == code:
                                chart_type_name = name
                                break

                        # 소스 데이터 주소
                        source_data = ""
                        try:
                            source_data = chart.SeriesCollection(1).Formula if chart.SeriesCollection().Count > 0 else ""
                        except:
                            pass

                        charts.append(
                            ChartInfo(
                                name=chart_obj.Name,
                                chart_type=chart_type_name,
                                source_data=source_data,
                                sheet_name=ws.Name,
                                left=chart_obj.Left,
                                top=chart_obj.Top,
                                width=chart_obj.Width,
                                height=chart_obj.Height,
                                has_title=chart.HasTitle,
                                title=chart.ChartTitle.Text if chart.HasTitle else None,
                            )
                        )
                except:
                    continue

            return charts

        except Exception as e:
            if sheet and "Subscript out of range" in str(e):
                raise SheetNotFoundError(sheet)
            raise COMError(f"차트 목록 조회 실패: {str(e)}")

    def configure_chart(self, workbook: Any, chart_name: str, **kwargs):
        """차트 설정"""
        try:
            # 차트 찾기
            chart_obj = None
            for ws in workbook.Sheets:
                try:
                    chart_obj = ws.ChartObjects(chart_name)
                    break
                except:
                    continue

            if not chart_obj:
                raise ChartNotFoundError(chart_name)

            chart = chart_obj.Chart

            # 제목 설정
            if "title" in kwargs:
                chart.HasTitle = True
                chart.ChartTitle.Text = kwargs["title"]

            # 범례 설정
            if "legend_position" in kwargs:
                legend_position = kwargs["legend_position"]
                legend_map = {
                    "top": -4160,
                    "bottom": -4107,
                    "left": -4131,
                    "right": -4152,
                }
                if legend_position.lower() in legend_map:
                    chart.HasLegend = True
                    chart.Legend.Position = legend_map[legend_position.lower()]
                elif legend_position.lower() == "none":
                    chart.HasLegend = False

            # 데이터 레이블
            if "show_data_labels" in kwargs:
                try:
                    for series in chart.SeriesCollection():
                        series.HasDataLabels = kwargs["show_data_labels"]
                except:
                    pass

            # 스타일
            if "style" in kwargs:
                try:
                    chart.ChartStyle = kwargs["style"]
                except:
                    pass

        except ChartNotFoundError:
            raise

        except Exception as e:
            raise COMError(f"차트 설정 실패: {str(e)}")

    def position_chart(self, workbook: Any, sheet: str, chart_name: str, left: int, top: int, width: int, height: int):
        """차트 위치 및 크기 조정"""
        try:
            ws = workbook.Sheets(sheet)
            chart_obj = ws.ChartObjects(chart_name)

            chart_obj.Left = left
            chart_obj.Top = top
            chart_obj.Width = width
            chart_obj.Height = height

        except Exception as e:
            if "Subscript out of range" in str(e):
                if sheet not in [s.Name for s in workbook.Sheets]:
                    raise SheetNotFoundError(sheet)
                else:
                    raise ChartNotFoundError(chart_name)
            raise COMError(f"차트 위치 조정 실패: {str(e)}")

    def export_chart(self, workbook: Any, sheet: str, chart_name: str, output_path: str, image_format: str = "png"):
        """차트 이미지 내보내기"""
        try:
            ws = workbook.Sheets(sheet)
            chart_obj = ws.ChartObjects(chart_name)

            # 절대 경로로 변환
            abs_path = str(Path(output_path).resolve())

            # 차트 내보내기
            chart_obj.Chart.Export(abs_path)

        except Exception as e:
            if "Subscript out of range" in str(e):
                if sheet not in [s.Name for s in workbook.Sheets]:
                    raise SheetNotFoundError(sheet)
                else:
                    raise ChartNotFoundError(chart_name)
            raise COMError(f"차트 내보내기 실패: {str(e)}")

    def delete_chart(self, workbook: Any, sheet: str, chart_name: str):
        """차트 삭제"""
        try:
            ws = workbook.Sheets(sheet)
            chart_obj = ws.ChartObjects(chart_name)
            chart_obj.Delete()

        except Exception as e:
            if "Subscript out of range" in str(e):
                if sheet not in [s.Name for s in workbook.Sheets]:
                    raise SheetNotFoundError(sheet)
                else:
                    raise ChartNotFoundError(chart_name)
            raise COMError(f"차트 삭제 실패: {str(e)}")

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
        """피벗 차트 생성 (Windows 전용)"""
        try:
            # 피벗테이블 생성 (간단한 구현)
            source_ws = workbook.Sheets(source_sheet)
            dest_ws = workbook.Sheets(dest_sheet)

            # 피벗 캐시 생성
            source_data = source_ws.Range(source_range)
            pivot_cache = workbook.PivotCaches().Create(SourceType=self._constants.xlDatabase, SourceData=source_data)

            # 피벗테이블 생성
            pivot_table = pivot_cache.CreatePivotTable(
                TableDestination=dest_ws.Range(dest_range), TableName=f"PivotTable_{dest_sheet}"
            )

            # 피벗 차트 생성
            chart_type_code = self.CHART_TYPE_MAP.get(chart_type.lower(), 51)
            pivot_chart = dest_ws.Shapes.AddChart2().Chart
            pivot_chart.SetSourceData(pivot_table.TableRange2)
            pivot_chart.ChartType = chart_type_code

            return pivot_chart.Parent.Name

        except Exception as e:
            raise COMError(f"피벗 차트 생성 실패: {str(e)}")

    # ===========================================
    # 헬퍼 메서드
    # ===========================================

    def _release_com_object(self, obj):
        """COM 객체 명시적 해제"""
        try:
            del obj
            gc.collect()
        except:
            pass

    # ===========================================
    # 헬퍼 메서드 (워크북 객체 접근)
    # ===========================================

    def get_active_workbook(self) -> Any:
        """활성 워크북 COM 객체 반환"""
        try:
            if self.xl.Workbooks.Count == 0:
                raise WorkbookNotFoundError("열린 워크북이 없습니다")
            return self.xl.ActiveWorkbook
        except Exception as e:
            raise COMError(f"활성 워크북 가져오기 실패: {str(e)}")

    def get_workbook_by_name(self, name: str) -> Any:
        """이름으로 워크북 COM 객체 찾기"""
        try:
            for wb in self.xl.Workbooks:
                if wb.Name == name:
                    return wb
            raise WorkbookNotFoundError(f"워크북 '{name}'을 찾을 수 없습니다")
        except WorkbookNotFoundError:
            raise
        except Exception as e:
            raise COMError(f"워크북 찾기 실패: {str(e)}")
