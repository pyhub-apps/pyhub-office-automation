"""
Excel 자동화 엔진 추상 인터페이스

이 모듈은 플랫폼별 Excel 자동화 엔진의 추상 기반 클래스를 정의합니다.
Windows(pywin32 COM)와 macOS(AppleScript) 구현체가 이 인터페이스를 따릅니다.
"""

from abc import ABC, abstractmethod
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple, Union


@dataclass
class WorkbookInfo:
    """워크북 정보 데이터 클래스"""

    name: str
    saved: bool
    full_name: str
    sheet_count: int
    active_sheet: str
    file_size_bytes: Optional[int] = None
    last_modified: Optional[str] = None


@dataclass
class RangeData:
    """셀 범위 데이터 클래스"""

    values: Any
    formulas: Optional[Any]
    address: str
    sheet_name: str
    row_count: int = 1
    column_count: int = 1
    cells_count: int = 1


@dataclass
class TableInfo:
    """테이블 정보 데이터 클래스"""

    name: str
    sheet_name: str
    address: str
    row_count: int
    column_count: int
    headers: List[str]
    sample_data: Optional[List[List[Any]]] = None


@dataclass
class ChartInfo:
    """차트 정보 데이터 클래스"""

    name: str
    chart_type: str
    source_data: str
    sheet_name: str
    left: float
    top: float
    width: float
    height: float
    has_title: bool = False
    title: Optional[str] = None


@dataclass
class PivotTableInfo:
    """피벗테이블 정보 데이터 클래스"""

    name: str
    sheet_name: str
    source_data: str
    row_fields: List[str]
    column_fields: List[str]
    value_fields: List[str]
    filter_fields: List[str]


@dataclass
class SlicerInfo:
    """슬라이서 정보 데이터 클래스"""

    name: str
    sheet_name: str
    caption: str
    source_field: str
    left: float
    top: float
    width: float
    height: float


@dataclass
class ShapeInfo:
    """도형 정보 데이터 클래스"""

    name: str
    sheet_name: str
    shape_type: str
    left: float
    top: float
    width: float
    height: float
    has_text: bool = False
    text: Optional[str] = None


class ExcelEngineBase(ABC):
    """
    Excel 자동화 엔진 추상 기반 클래스

    모든 플랫폼별 구현체(WindowsEngine, MacOSEngine)는 이 클래스를 상속받아
    43개 Excel 명령어에 대응하는 메서드를 구현해야 합니다.

    Issue #87: 핵심 22개 명령어 (완료)
    Issue #88: 추가 21개 명령어 (진행 중)
    """

    # ===========================================
    # 워크북 관리 (4개 명령어)
    # ===========================================

    @abstractmethod
    def get_workbooks(self) -> List[WorkbookInfo]:
        """
        현재 열려있는 모든 워크북 목록을 조회합니다.

        Returns:
            List[WorkbookInfo]: 워크북 정보 리스트

        Raises:
            RuntimeError: Excel이 실행되지 않았거나 접근 불가능한 경우

        CLI 명령어: workbook-list
        """
        pass

    @abstractmethod
    def get_workbook_info(self, workbook: Any) -> Dict[str, Any]:
        """
        특정 워크북의 상세 정보를 조회합니다.

        Args:
            workbook: 워크북 객체 (플랫폼별로 다름)

        Returns:
            Dict[str, Any]: 워크북 상세 정보

        CLI 명령어: workbook-info
        """
        pass

    @abstractmethod
    def open_workbook(self, file_path: str, visible: bool = False) -> Any:
        """
        Excel 워크북을 엽니다.

        Args:
            file_path: 워크북 파일 경로
            visible: Excel 애플리케이션 표시 여부

        Returns:
            Any: 워크북 객체 (플랫폼별로 다름)

        Raises:
            FileNotFoundError: 파일이 존재하지 않는 경우
            RuntimeError: 워크북을 열 수 없는 경우

        CLI 명령어: workbook-open
        """
        pass

    @abstractmethod
    def create_workbook(self, save_path: Optional[str] = None, visible: bool = False) -> Any:
        """
        새 워크북을 생성합니다.

        Args:
            save_path: 저장할 파일 경로 (None이면 저장하지 않음)
            visible: Excel 애플리케이션 표시 여부

        Returns:
            Any: 워크북 객체

        CLI 명령어: workbook-create
        """
        pass

    # ===========================================
    # 시트 관리 (4개 명령어)
    # ===========================================

    @abstractmethod
    def activate_sheet(self, workbook: Any, sheet_name: str):
        """
        시트를 활성화합니다.

        Args:
            workbook: 워크북 객체
            sheet_name: 활성화할 시트 이름

        Raises:
            ValueError: 시트가 존재하지 않는 경우

        CLI 명령어: sheet-activate
        """
        pass

    @abstractmethod
    def add_sheet(self, workbook: Any, name: str, before: Optional[str] = None) -> str:
        """
        새 시트를 추가합니다.

        Args:
            workbook: 워크북 객체
            name: 새 시트 이름
            before: 이 시트 앞에 추가 (None이면 마지막에 추가)

        Returns:
            str: 생성된 시트 이름

        Raises:
            ValueError: 동일한 이름의 시트가 이미 존재하는 경우

        CLI 명령어: sheet-add
        """
        pass

    @abstractmethod
    def delete_sheet(self, workbook: Any, sheet_name: str):
        """
        시트를 삭제합니다.

        Args:
            workbook: 워크북 객체
            sheet_name: 삭제할 시트 이름

        Raises:
            ValueError: 시트가 존재하지 않는 경우
            RuntimeError: 마지막 시트는 삭제할 수 없음

        CLI 명령어: sheet-delete
        """
        pass

    @abstractmethod
    def rename_sheet(self, workbook: Any, old_name: str, new_name: str):
        """
        시트 이름을 변경합니다.

        Args:
            workbook: 워크북 객체
            old_name: 기존 시트 이름
            new_name: 새 시트 이름

        Raises:
            ValueError: 시트가 존재하지 않거나 새 이름이 이미 존재하는 경우

        CLI 명령어: sheet-rename
        """
        pass

    # ===========================================
    # 데이터 읽기/쓰기 (2개 명령어)
    # ===========================================

    @abstractmethod
    def read_range(
        self, workbook: Any, sheet: str, range_str: str, expand: Optional[str] = None, include_formulas: bool = True
    ) -> RangeData:
        """
        셀 범위 데이터를 읽습니다.

        Args:
            workbook: 워크북 객체
            sheet: 시트 이름
            range_str: 범위 문자열 (예: "A1:C10")
            expand: 범위 확장 모드 ("table", "down", "right", None)
            include_formulas: 공식 포함 여부

        Returns:
            RangeData: 범위 데이터

        CLI 명령어: range-read
        """
        pass

    @abstractmethod
    def write_range(self, workbook: Any, sheet: str, range_str: str, data: Any, include_formulas: bool = False):
        """
        셀 범위에 데이터를 씁니다.

        Args:
            workbook: 워크북 객체
            sheet: 시트 이름
            range_str: 시작 셀 주소 (예: "A1")
            data: 쓸 데이터 (단일값, 1차원 리스트, 2차원 리스트)
            include_formulas: 공식 포함 여부

        CLI 명령어: range-write
        """
        pass

    # ===========================================
    # 테이블 (5개 명령어)
    # ===========================================

    @abstractmethod
    def list_tables(self, workbook: Any, sheet: Optional[str] = None) -> List[TableInfo]:
        """
        워크북의 테이블 목록을 조회합니다.

        Args:
            workbook: 워크북 객체
            sheet: 시트 이름 (None이면 전체 시트)

        Returns:
            List[TableInfo]: 테이블 정보 리스트

        CLI 명령어: table-list
        """
        pass

    @abstractmethod
    def read_table(
        self,
        workbook: Any,
        table_name: str,
        columns: Optional[List[str]] = None,
        limit: Optional[int] = None,
        offset: int = 0,
    ) -> Dict[str, Any]:
        """
        테이블 데이터를 읽습니다.

        Args:
            workbook: 워크북 객체
            table_name: 테이블 이름
            columns: 읽을 컬럼 리스트 (None이면 전체)
            limit: 읽을 행 개수 제한
            offset: 시작 행 오프셋

        Returns:
            Dict[str, Any]: 테이블 데이터

        CLI 명령어: table-read
        """
        pass

    @abstractmethod
    def write_table(self, workbook: Any, sheet: str, table_name: str, data: List[List[Any]], start_cell: str = "A1"):
        """
        테이블에 데이터를 씁니다.

        Args:
            workbook: 워크북 객체
            sheet: 시트 이름
            table_name: 테이블 이름
            data: 쓸 데이터 (헤더 포함)
            start_cell: 시작 셀 주소

        CLI 명령어: table-write
        """
        pass

    @abstractmethod
    def analyze_table(self, workbook: Any, table_name: str) -> Dict[str, Any]:
        """
        테이블 데이터를 분석합니다.

        Args:
            workbook: 워크북 객체
            table_name: 테이블 이름

        Returns:
            Dict[str, Any]: 분석 결과 (통계, 데이터 타입 등)

        CLI 명령어: table-analyze
        """
        pass

    @abstractmethod
    def generate_metadata(self, workbook: Any) -> Dict[str, Any]:
        """
        워크북의 메타데이터를 생성합니다.

        Args:
            workbook: 워크북 객체

        Returns:
            Dict[str, Any]: 메타데이터

        CLI 명령어: metadata-generate
        """
        pass

    # ===========================================
    # 차트 (7개 명령어)
    # ===========================================

    @abstractmethod
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
        """
        차트를 생성합니다.

        Args:
            workbook: 워크북 객체
            sheet: 시트 이름
            data_range: 데이터 범위
            chart_type: 차트 타입 (column, bar, line, pie 등)
            position: 차트 위치 (셀 주소)
            width: 차트 너비 (픽셀)
            height: 차트 높이 (픽셀)
            title: 차트 제목
            **kwargs: 추가 옵션

        Returns:
            str: 생성된 차트 이름

        CLI 명령어: chart-add
        """
        pass

    @abstractmethod
    def list_charts(self, workbook: Any, sheet: Optional[str] = None) -> List[ChartInfo]:
        """
        차트 목록을 조회합니다.

        Args:
            workbook: 워크북 객체
            sheet: 시트 이름 (None이면 전체 시트)

        Returns:
            List[ChartInfo]: 차트 정보 리스트

        CLI 명령어: chart-list
        """
        pass

    @abstractmethod
    def configure_chart(self, workbook: Any, chart_name: str, **kwargs):
        """
        차트를 설정합니다.

        Args:
            workbook: 워크북 객체
            chart_name: 차트 이름
            **kwargs: 설정 옵션 (title, legend_position, show_data_labels 등)

        CLI 명령어: chart-configure
        """
        pass

    @abstractmethod
    def position_chart(self, workbook: Any, sheet: str, chart_name: str, left: int, top: int, width: int, height: int):
        """
        차트 위치와 크기를 조정합니다.

        Args:
            workbook: 워크북 객체
            sheet: 시트 이름
            chart_name: 차트 이름
            left: 왼쪽 위치 (픽셀)
            top: 상단 위치 (픽셀)
            width: 너비 (픽셀)
            height: 높이 (픽셀)

        CLI 명령어: chart-position
        """
        pass

    @abstractmethod
    def export_chart(self, workbook: Any, sheet: str, chart_name: str, output_path: str, image_format: str = "png"):
        """
        차트를 이미지로 내보냅니다.

        Args:
            workbook: 워크북 객체
            sheet: 시트 이름
            chart_name: 차트 이름
            output_path: 출력 파일 경로
            image_format: 이미지 포맷 (png, jpg 등)

        CLI 명령어: chart-export
        """
        pass

    @abstractmethod
    def delete_chart(self, workbook: Any, sheet: str, chart_name: str):
        """
        차트를 삭제합니다.

        Args:
            workbook: 워크북 객체
            sheet: 시트 이름
            chart_name: 차트 이름

        CLI 명령어: chart-delete
        """
        pass

    @abstractmethod
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
        """
        피벗 차트를 생성합니다. (Windows 우선 지원)

        Args:
            workbook: 워크북 객체
            source_sheet: 원본 데이터 시트
            source_range: 원본 데이터 범위
            dest_sheet: 대상 시트
            dest_range: 대상 범위
            chart_type: 차트 타입
            **kwargs: 추가 옵션

        Returns:
            str: 생성된 피벗 차트 이름

        Raises:
            NotImplementedError: macOS에서 완전히 지원되지 않을 수 있음

        CLI 명령어: chart-pivot-create
        """
        pass

    # ===========================================
    # 헬퍼 메서드 (서브클래스에서 선택적 구현)
    # ===========================================

    @abstractmethod
    def get_active_workbook(self) -> Any:
        """
        현재 활성 워크북 객체를 반환합니다.

        Returns:
            Any: 활성 워크북 객체 (플랫폼별로 다름)
                 - Windows: COM Workbook 객체
                 - macOS: 워크북 이름 (str)

        Raises:
            WorkbookNotFoundError: 열린 워크북이 없는 경우
        """
        pass

    @abstractmethod
    def get_workbook_by_name(self, name: str) -> Any:
        """
        이름으로 워크북을 찾아 반환합니다.

        Args:
            name: 워크북 이름

        Returns:
            Any: 워크북 객체 (플랫폼별로 다름)
                 - Windows: COM Workbook 객체
                 - macOS: 워크북 이름 (str)

        Raises:
            WorkbookNotFoundError: 워크북을 찾을 수 없는 경우
        """
        pass

    # ===========================================
    # 피벗 테이블 (5개 명령어) - Issue #88
    # ===========================================

    @abstractmethod
    def create_pivot_table(
        self,
        workbook: Any,
        source_sheet: str,
        source_range: str,
        dest_sheet: str,
        dest_cell: str,
        pivot_name: Optional[str] = None,
        **kwargs,
    ) -> Dict[str, Any]:
        """
        피벗 테이블을 생성합니다.

        Args:
            workbook: 워크북 객체
            source_sheet: 원본 데이터 시트 이름
            source_range: 원본 데이터 범위 (예: "A1:D100")
            dest_sheet: 피벗 테이블을 생성할 시트 이름
            dest_cell: 피벗 테이블 시작 위치 (예: "F1")
            pivot_name: 피벗 테이블 이름 (None이면 자동 생성)
            **kwargs: 추가 옵션

        Returns:
            Dict[str, Any]: 생성된 피벗 테이블 정보

        Raises:
            EngineNotSupportedError: macOS에서 지원하지 않는 경우

        CLI 명령어: pivot-create
        """
        pass

    @abstractmethod
    def configure_pivot_table(
        self,
        workbook: Any,
        sheet: str,
        pivot_name: str,
        row_fields: Optional[List[str]] = None,
        column_fields: Optional[List[str]] = None,
        value_fields: Optional[List[Tuple[str, str]]] = None,
        filter_fields: Optional[List[str]] = None,
        **kwargs,
    ):
        """
        피벗 테이블을 설정합니다.

        Args:
            workbook: 워크북 객체
            sheet: 시트 이름
            pivot_name: 피벗 테이블 이름
            row_fields: 행 필드 리스트
            column_fields: 열 필드 리스트
            value_fields: 값 필드 리스트 [(필드명, 집계함수), ...]
            filter_fields: 필터 필드 리스트
            **kwargs: 추가 옵션

        Raises:
            EngineNotSupportedError: macOS에서 지원하지 않는 경우

        CLI 명령어: pivot-configure
        """
        pass

    @abstractmethod
    def refresh_pivot_table(self, workbook: Any, sheet: str, pivot_name: str):
        """
        피벗 테이블 데이터를 새로고침합니다.

        Args:
            workbook: 워크북 객체
            sheet: 시트 이름
            pivot_name: 피벗 테이블 이름

        CLI 명령어: pivot-refresh
        """
        pass

    @abstractmethod
    def delete_pivot_table(self, workbook: Any, sheet: str, pivot_name: str):
        """
        피벗 테이블을 삭제합니다.

        Args:
            workbook: 워크북 객체
            sheet: 시트 이름
            pivot_name: 피벗 테이블 이름

        CLI 명령어: pivot-delete
        """
        pass

    @abstractmethod
    def list_pivot_tables(self, workbook: Any, sheet: Optional[str] = None) -> List[PivotTableInfo]:
        """
        피벗 테이블 목록을 조회합니다.

        Args:
            workbook: 워크북 객체
            sheet: 시트 이름 (None이면 전체 시트)

        Returns:
            List[PivotTableInfo]: 피벗 테이블 정보 리스트

        CLI 명령어: pivot-list
        """
        pass

    # ===========================================
    # 슬라이서 (4개 명령어) - Issue #88
    # ===========================================

    @abstractmethod
    def add_slicer(
        self,
        workbook: Any,
        sheet: str,
        pivot_name: str,
        field_name: str,
        left: int,
        top: int,
        width: int = 200,
        height: int = 150,
        slicer_name: Optional[str] = None,
        **kwargs,
    ) -> Dict[str, Any]:
        """
        슬라이서를 추가합니다. (Windows 전용)

        Args:
            workbook: 워크북 객체
            sheet: 시트 이름
            pivot_name: 피벗 테이블 이름
            field_name: 슬라이서 필드 이름
            left: 왼쪽 위치 (픽셀)
            top: 상단 위치 (픽셀)
            width: 너비 (픽셀)
            height: 높이 (픽셀)
            slicer_name: 슬라이서 이름 (None이면 자동 생성)
            **kwargs: 추가 옵션 (caption, style, columns 등)

        Returns:
            Dict[str, Any]: 생성된 슬라이서 정보

        Raises:
            EngineNotSupportedError: macOS에서는 지원하지 않음

        CLI 명령어: slicer-add
        """
        pass

    @abstractmethod
    def list_slicers(self, workbook: Any, sheet: Optional[str] = None) -> List[SlicerInfo]:
        """
        슬라이서 목록을 조회합니다.

        Args:
            workbook: 워크북 객체
            sheet: 시트 이름 (None이면 전체 시트)

        Returns:
            List[SlicerInfo]: 슬라이서 정보 리스트

        Raises:
            EngineNotSupportedError: macOS에서는 지원하지 않음

        CLI 명령어: slicer-list
        """
        pass

    @abstractmethod
    def position_slicer(
        self,
        workbook: Any,
        sheet: str,
        slicer_name: str,
        left: int,
        top: int,
        width: Optional[int] = None,
        height: Optional[int] = None,
    ):
        """
        슬라이서 위치와 크기를 조정합니다.

        Args:
            workbook: 워크북 객체
            sheet: 시트 이름
            slicer_name: 슬라이서 이름
            left: 왼쪽 위치 (픽셀)
            top: 상단 위치 (픽셀)
            width: 너비 (픽셀, None이면 유지)
            height: 높이 (픽셀, None이면 유지)

        Raises:
            EngineNotSupportedError: macOS에서는 지원하지 않음

        CLI 명령어: slicer-position
        """
        pass

    @abstractmethod
    def connect_slicer(self, workbook: Any, slicer_name: str, pivot_names: List[str]):
        """
        슬라이서를 여러 피벗 테이블에 연결합니다.

        Args:
            workbook: 워크북 객체
            slicer_name: 슬라이서 이름
            pivot_names: 연결할 피벗 테이블 이름 리스트

        Raises:
            EngineNotSupportedError: macOS에서는 지원하지 않음

        CLI 명령어: slicer-connect
        """
        pass

    # ===========================================
    # 도형 (5개 명령어) - Issue #88
    # ===========================================

    @abstractmethod
    def add_shape(
        self,
        workbook: Any,
        sheet: str,
        shape_type: str,
        left: int,
        top: int,
        width: int,
        height: int,
        shape_name: Optional[str] = None,
        **kwargs,
    ) -> Dict[str, Any]:
        """
        도형을 추가합니다.

        Args:
            workbook: 워크북 객체
            sheet: 시트 이름
            shape_type: 도형 유형 (rectangle, oval, line, arrow 등)
            left: 왼쪽 위치 (픽셀)
            top: 상단 위치 (픽셀)
            width: 너비 (픽셀)
            height: 높이 (픽셀)
            shape_name: 도형 이름 (None이면 자동 생성)
            **kwargs: 추가 옵션 (fill_color, transparency 등)

        Returns:
            Dict[str, Any]: 생성된 도형 정보

        Note:
            고급 기능(뉴모피즘 스타일 등)은 xlwings 하이브리드 사용

        CLI 명령어: shape-add
        """
        pass

    @abstractmethod
    def delete_shape(self, workbook: Any, sheet: str, shape_name: str):
        """
        도형을 삭제합니다.

        Args:
            workbook: 워크북 객체
            sheet: 시트 이름
            shape_name: 도형 이름

        CLI 명령어: shape-delete
        """
        pass

    @abstractmethod
    def list_shapes(self, workbook: Any, sheet: str) -> List[ShapeInfo]:
        """
        도형 목록을 조회합니다.

        Args:
            workbook: 워크북 객체
            sheet: 시트 이름

        Returns:
            List[ShapeInfo]: 도형 정보 리스트

        CLI 명령어: shape-list
        """
        pass

    @abstractmethod
    def format_shape(self, workbook: Any, sheet: str, shape_name: str, **kwargs):
        """
        도형 서식을 설정합니다.

        Args:
            workbook: 워크북 객체
            sheet: 시트 이름
            shape_name: 도형 이름
            **kwargs: 서식 옵션 (fill_color, line_color, line_width 등)

        Note:
            복잡한 서식은 xlwings 하이브리드 사용 권장

        CLI 명령어: shape-format
        """
        pass

    @abstractmethod
    def group_shapes(self, workbook: Any, sheet: str, shape_names: List[str], group_name: Optional[str] = None) -> str:
        """
        여러 도형을 그룹화합니다.

        Args:
            workbook: 워크북 객체
            sheet: 시트 이름
            shape_names: 그룹화할 도형 이름 리스트
            group_name: 그룹 이름 (None이면 자동 생성)

        Returns:
            str: 생성된 그룹 이름

        Note:
            macOS에서는 제한적 지원

        CLI 명령어: shape-group
        """
        pass

    # ===========================================
    # 테이블 추가 기능 (4개 명령어) - Issue #88
    # ===========================================

    @abstractmethod
    def create_table(
        self, workbook: Any, sheet: str, range_str: str, table_name: Optional[str] = None, has_headers: bool = True, **kwargs
    ) -> Dict[str, Any]:
        """
        Excel 테이블(ListObject)을 생성합니다.

        Args:
            workbook: 워크북 객체
            sheet: 시트 이름
            range_str: 테이블 범위 (예: "A1:D100")
            table_name: 테이블 이름 (None이면 자동 생성)
            has_headers: 첫 행이 헤더인지 여부
            **kwargs: 추가 옵션 (table_style 등)

        Returns:
            Dict[str, Any]: 생성된 테이블 정보

        CLI 명령어: table-create
        """
        pass

    @abstractmethod
    def sort_table(self, workbook: Any, sheet: str, table_name: str, sort_fields: List[Tuple[str, str]]):
        """
        테이블을 정렬합니다.

        Args:
            workbook: 워크북 객체
            sheet: 시트 이름
            table_name: 테이블 이름
            sort_fields: 정렬 필드 리스트 [(컬럼명, 정렬방향), ...]
                        정렬방향: "asc" 또는 "desc"

        CLI 명령어: table-sort
        """
        pass

    @abstractmethod
    def clear_table_sort(self, workbook: Any, sheet: str, table_name: str):
        """
        테이블 정렬을 해제합니다.

        Args:
            workbook: 워크북 객체
            sheet: 시트 이름
            table_name: 테이블 이름

        CLI 명령어: table-sort-clear
        """
        pass

    @abstractmethod
    def get_table_sort_info(self, workbook: Any, sheet: str, table_name: str) -> Dict[str, Any]:
        """
        테이블 정렬 정보를 조회합니다.

        Args:
            workbook: 워크북 객체
            sheet: 시트 이름
            table_name: 테이블 이름

        Returns:
            Dict[str, Any]: 정렬 정보

        CLI 명령어: table-sort-info
        """
        pass

    # ===========================================
    # 데이터 변환 (3개 명령어) - Issue #88
    # ===========================================

    @abstractmethod
    def analyze_data(self, workbook: Any, sheet: str, range_str: str, **kwargs) -> Dict[str, Any]:
        """
        데이터를 분석합니다.

        Args:
            workbook: 워크북 객체
            sheet: 시트 이름
            range_str: 분석할 범위
            **kwargs: 분석 옵션

        Returns:
            Dict[str, Any]: 분석 결과 (통계, 데이터 타입 등)

        CLI 명령어: data-analyze
        """
        pass

    @abstractmethod
    def transform_data(self, workbook: Any, sheet: str, range_str: str, transform_type: str, **kwargs):
        """
        데이터를 변환합니다.

        Args:
            workbook: 워크북 객체
            sheet: 시트 이름
            range_str: 변환할 범위
            transform_type: 변환 유형 (transpose, normalize 등)
            **kwargs: 변환 옵션

        CLI 명령어: data-transform
        """
        pass

    @abstractmethod
    def convert_range(self, workbook: Any, sheet: str, range_str: str, target_type: str, **kwargs):
        """
        셀 범위의 데이터 형식을 변환합니다.

        Args:
            workbook: 워크북 객체
            sheet: 시트 이름
            range_str: 변환할 범위
            target_type: 목표 데이터 타입 (number, text, date 등)
            **kwargs: 변환 옵션

        CLI 명령어: range-convert
        """
        pass
