"""
Engine Factory 테스트

플랫폼 자동 감지 및 올바른 엔진 반환을 검증합니다.
"""

import platform

import pytest


class TestEngineFactory:
    """Engine Factory 기능 테스트"""

    def test_get_platform_name(self):
        """플랫폼 이름 감지"""
        from pyhub_office_automation.excel.engines import get_platform_name

        platform_name = get_platform_name()

        assert platform_name in ["Windows", "macOS", "Unknown"]

        # 현재 플랫폼 확인
        if platform.system() == "Windows":
            assert platform_name == "Windows"
        elif platform.system() == "Darwin":
            assert platform_name == "macOS"

    def test_get_engine_auto_detect(self):
        """플랫폼 자동 감지로 엔진 획득"""
        from pyhub_office_automation.excel.engines import get_engine

        engine = get_engine()

        assert engine is not None

        # 플랫폼별 엔진 타입 확인
        if platform.system() == "Windows":
            from pyhub_office_automation.excel.engines.windows import WindowsEngine

            assert isinstance(engine, WindowsEngine)
        elif platform.system() == "Darwin":
            from pyhub_office_automation.excel.engines.macos import MacOSEngine

            assert isinstance(engine, MacOSEngine)

    def test_get_engine_singleton(self):
        """엔진 싱글톤 패턴 검증"""
        from pyhub_office_automation.excel.engines import get_engine

        engine1 = get_engine()
        engine2 = get_engine()

        assert engine1 is engine2

    def test_reset_engine(self):
        """엔진 리셋 기능"""
        from pyhub_office_automation.excel.engines import get_engine, reset_engine

        engine1 = get_engine()

        reset_engine()

        engine2 = get_engine()

        # 리셋 후 다른 인스턴스여야 함
        assert engine1 is not engine2

    @pytest.mark.skipif(platform.system() != "Windows", reason="Windows only")
    def test_force_windows_engine(self):
        """Windows 엔진 강제 지정"""
        from pyhub_office_automation.excel.engines import get_engine, reset_engine
        from pyhub_office_automation.excel.engines.windows import WindowsEngine

        reset_engine()
        engine = get_engine(force_platform="Windows")

        assert isinstance(engine, WindowsEngine)

    @pytest.mark.skipif(platform.system() == "Windows", reason="Non-Windows only")
    def test_unsupported_platform_error(self):
        """지원되지 않는 플랫폼 오류"""
        from pyhub_office_automation.excel.engines import PlatformNotSupportedError, get_engine, reset_engine

        reset_engine()

        with pytest.raises(PlatformNotSupportedError):
            get_engine(force_platform="Linux")


class TestEngineInterface:
    """엔진 인터페이스 일관성 테스트"""

    def test_engine_has_all_methods(self):
        """엔진이 모든 필수 메서드를 가지고 있는지 확인"""
        from pyhub_office_automation.excel.engines import get_engine

        engine = get_engine()

        # 워크북 관리 (4)
        assert hasattr(engine, "get_workbooks")
        assert hasattr(engine, "get_workbook_info")
        assert hasattr(engine, "open_workbook")
        assert hasattr(engine, "create_workbook")

        # 시트 관리 (4)
        assert hasattr(engine, "activate_sheet")
        assert hasattr(engine, "add_sheet")
        assert hasattr(engine, "delete_sheet")
        assert hasattr(engine, "rename_sheet")

        # 데이터 (2)
        assert hasattr(engine, "read_range")
        assert hasattr(engine, "write_range")

        # 테이블 (5)
        assert hasattr(engine, "list_tables")
        assert hasattr(engine, "read_table")
        assert hasattr(engine, "write_table")
        assert hasattr(engine, "analyze_table")
        assert hasattr(engine, "generate_metadata")

        # 차트 (7)
        assert hasattr(engine, "add_chart")
        assert hasattr(engine, "list_charts")
        assert hasattr(engine, "configure_chart")
        assert hasattr(engine, "position_chart")
        assert hasattr(engine, "export_chart")
        assert hasattr(engine, "delete_chart")
        assert hasattr(engine, "create_pivot_chart")

    def test_engine_methods_callable(self):
        """엔진 메서드가 호출 가능한지 확인"""
        from pyhub_office_automation.excel.engines import get_engine

        engine = get_engine()

        # 모든 메서드가 callable인지 확인
        methods = [
            "get_workbooks",
            "get_workbook_info",
            "open_workbook",
            "create_workbook",
            "activate_sheet",
            "add_sheet",
            "delete_sheet",
            "rename_sheet",
            "read_range",
            "write_range",
            "list_tables",
            "read_table",
            "write_table",
            "analyze_table",
            "generate_metadata",
            "add_chart",
            "list_charts",
            "configure_chart",
            "position_chart",
            "export_chart",
            "delete_chart",
            "create_pivot_chart",
        ]

        for method_name in methods:
            method = getattr(engine, method_name)
            assert callable(method), f"{method_name} is not callable"


class TestDataClasses:
    """데이터 클래스 테스트"""

    def test_workbook_info_dataclass(self):
        """WorkbookInfo 데이터 클래스"""
        from pyhub_office_automation.excel.engines import WorkbookInfo

        wb_info = WorkbookInfo(
            name="test.xlsx", saved=True, full_name="/path/to/test.xlsx", sheet_count=3, active_sheet="Sheet1"
        )

        assert wb_info.name == "test.xlsx"
        assert wb_info.saved is True
        assert wb_info.sheet_count == 3
        assert wb_info.active_sheet == "Sheet1"

    def test_range_data_dataclass(self):
        """RangeData 데이터 클래스"""
        from pyhub_office_automation.excel.engines import RangeData

        range_data = RangeData(
            values=[[1, 2], [3, 4]],
            formulas=None,
            address="$A$1:$B$2",
            sheet_name="Sheet1",
            row_count=2,
            column_count=2,
            cells_count=4,
        )

        assert range_data.values == [[1, 2], [3, 4]]
        assert range_data.address == "$A$1:$B$2"
        assert range_data.row_count == 2
        assert range_data.cells_count == 4

    def test_table_info_dataclass(self):
        """TableInfo 데이터 클래스"""
        from pyhub_office_automation.excel.engines import TableInfo

        table_info = TableInfo(
            name="Table1", sheet_name="Sheet1", address="A1:C10", row_count=9, column_count=3, headers=["A", "B", "C"]
        )

        assert table_info.name == "Table1"
        assert table_info.row_count == 9
        assert len(table_info.headers) == 3

    def test_chart_info_dataclass(self):
        """ChartInfo 데이터 클래스"""
        from pyhub_office_automation.excel.engines import ChartInfo

        chart_info = ChartInfo(
            name="Chart1",
            chart_type="column",
            source_data="A1:B10",
            sheet_name="Sheet1",
            left=100,
            top=50,
            width=400,
            height=300,
            has_title=True,
            title="Test Chart",
        )

        assert chart_info.name == "Chart1"
        assert chart_info.chart_type == "column"
        assert chart_info.has_title is True
        assert chart_info.title == "Test Chart"


class TestExceptions:
    """예외 클래스 테스트"""

    def test_excel_engine_error(self):
        """ExcelEngineError 기본 예외"""
        from pyhub_office_automation.excel.engines import ExcelEngineError

        error = ExcelEngineError("Test error")
        assert str(error) == "Test error"

    def test_workbook_not_found_error(self):
        """WorkbookNotFoundError"""
        from pyhub_office_automation.excel.engines import WorkbookNotFoundError

        error = WorkbookNotFoundError("test.xlsx")
        assert error.workbook_name == "test.xlsx"
        assert "test.xlsx" in str(error)

    def test_sheet_not_found_error(self):
        """SheetNotFoundError"""
        from pyhub_office_automation.excel.engines import SheetNotFoundError

        error = SheetNotFoundError("Sheet1")
        assert error.sheet_name == "Sheet1"
        assert "Sheet1" in str(error)

    def test_table_not_found_error(self):
        """TableNotFoundError"""
        from pyhub_office_automation.excel.engines import TableNotFoundError

        error = TableNotFoundError("Table1")
        assert error.table_name == "Table1"
        assert "Table1" in str(error)

    def test_chart_not_found_error(self):
        """ChartNotFoundError"""
        from pyhub_office_automation.excel.engines import ChartNotFoundError

        error = ChartNotFoundError("Chart1")
        assert error.chart_name == "Chart1"
        assert "Chart1" in str(error)

    def test_platform_not_supported_error(self):
        """PlatformNotSupportedError"""
        from pyhub_office_automation.excel.engines import PlatformNotSupportedError

        error = PlatformNotSupportedError("Linux")
        assert error.platform == "Linux"
        assert "Linux" in str(error)

        error_with_feature = PlatformNotSupportedError("Linux", "pivot chart")
        assert error_with_feature.feature == "pivot chart"
        assert "pivot chart" in str(error_with_feature)


if __name__ == "__main__":
    pytest.main([__file__, "-v", "--tb=short"])
