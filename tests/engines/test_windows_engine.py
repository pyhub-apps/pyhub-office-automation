"""
WindowsEngine 통합 테스트

pywin32 COM 기반 WindowsEngine의 모든 기능을 검증합니다.
실제 Excel이 설치된 Windows 환경에서만 실행됩니다.
"""

import os
import platform
import tempfile
from pathlib import Path

import pytest

# Windows 환경에서만 테스트 실행
pytestmark = pytest.mark.skipif(platform.system() != "Windows", reason="Windows only")


@pytest.fixture
def engine():
    """WindowsEngine 인스턴스 생성"""
    from pyhub_office_automation.excel.engines import get_engine, reset_engine

    reset_engine()
    return get_engine(force_platform="Windows")


@pytest.fixture
def test_workbook(engine):
    """테스트용 워크북 생성"""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        wb = engine.create_workbook(save_path=tmp.name, visible=False)
        yield wb
        # 정리
        try:
            wb.Close(SaveChanges=False)
        except:
            pass
        try:
            os.unlink(tmp.name)
        except:
            pass


class TestWorkbookManagement:
    """워크북 관리 메서드 테스트"""

    def test_get_workbooks(self, engine, test_workbook):
        """워크북 목록 조회"""
        workbooks = engine.get_workbooks()
        assert len(workbooks) >= 1
        assert any(wb.name == test_workbook.Name for wb in workbooks)

    def test_get_workbook_info(self, engine, test_workbook):
        """워크북 상세 정보 조회"""
        info = engine.get_workbook_info(test_workbook)
        assert "name" in info
        assert "sheet_count" in info
        assert "sheets" in info
        assert info["sheet_count"] >= 1

    def test_open_workbook(self, engine):
        """워크북 열기"""
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            # 임시 워크북 생성
            wb1 = engine.create_workbook(save_path=tmp.name)
            wb1.Close(SaveChanges=True)

            # 워크북 열기
            wb2 = engine.open_workbook(tmp.name, visible=False)
            assert wb2.Name == Path(tmp.name).name

            # 정리
            wb2.Close(SaveChanges=False)
            os.unlink(tmp.name)

    def test_create_workbook(self, engine):
        """워크북 생성"""
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            wb = engine.create_workbook(save_path=tmp.name, visible=False)
            assert wb is not None
            assert os.path.exists(tmp.name)

            # 정리
            wb.Close(SaveChanges=False)
            os.unlink(tmp.name)


class TestSheetManagement:
    """시트 관리 메서드 테스트"""

    def test_activate_sheet(self, engine, test_workbook):
        """시트 활성화"""
        # 기본 시트 활성화
        engine.activate_sheet(test_workbook, "Sheet1")
        assert test_workbook.ActiveSheet.Name == "Sheet1"

    def test_add_sheet(self, engine, test_workbook):
        """시트 추가"""
        new_sheet_name = "TestSheet"
        engine.add_sheet(test_workbook, new_sheet_name)

        sheet_names = [sheet.Name for sheet in test_workbook.Sheets]
        assert new_sheet_name in sheet_names

    def test_delete_sheet(self, engine, test_workbook):
        """시트 삭제"""
        # 시트 추가 후 삭제
        engine.add_sheet(test_workbook, "ToDelete")
        engine.delete_sheet(test_workbook, "ToDelete")

        sheet_names = [sheet.Name for sheet in test_workbook.Sheets]
        assert "ToDelete" not in sheet_names

    def test_rename_sheet(self, engine, test_workbook):
        """시트 이름 변경"""
        old_name = "Sheet1"
        new_name = "RenamedSheet"

        engine.rename_sheet(test_workbook, old_name, new_name)

        sheet_names = [sheet.Name for sheet in test_workbook.Sheets]
        assert new_name in sheet_names
        assert old_name not in sheet_names


class TestDataOperations:
    """데이터 읽기/쓰기 테스트"""

    def test_write_and_read_range(self, engine, test_workbook):
        """범위 쓰기 및 읽기"""
        # 데이터 쓰기
        test_data = [["Name", "Age"], ["Alice", 30], ["Bob", 25]]
        engine.write_range(test_workbook, "Sheet1", "A1", test_data)

        # 데이터 읽기
        result = engine.read_range(test_workbook, "Sheet1", "A1:B3", include_formulas=False)

        assert result.values is not None
        assert result.row_count == 3
        assert result.column_count == 2

    def test_read_with_formula(self, engine, test_workbook):
        """공식 포함 읽기"""
        # 값과 공식 쓰기
        engine.write_range(test_workbook, "Sheet1", "A1", 10)
        engine.write_range(test_workbook, "Sheet1", "A2", 20)
        engine.write_range(test_workbook, "Sheet1", "A3", "=A1+A2", include_formulas=True)

        # 공식 포함 읽기
        result = engine.read_range(test_workbook, "Sheet1", "A3", include_formulas=True)

        assert result.values == 30
        assert result.formulas == "=A1+A2"

    def test_expand_range(self, engine, test_workbook):
        """범위 확장 모드 테스트"""
        # 테스트 데이터
        test_data = [["A", "B", "C"], [1, 2, 3], [4, 5, 6]]
        engine.write_range(test_workbook, "Sheet1", "A1", test_data)

        # table 확장
        result = engine.read_range(test_workbook, "Sheet1", "A1", expand="table", include_formulas=False)

        assert result.row_count == 3
        assert result.column_count == 3


class TestTableOperations:
    """테이블 메서드 테스트"""

    def test_write_and_list_tables(self, engine, test_workbook):
        """테이블 생성 및 목록 조회"""
        # 테이블 데이터 쓰기
        table_data = [["Product", "Price"], ["Apple", 1.2], ["Banana", 0.8]]
        engine.write_table(test_workbook, "Sheet1", "ProductTable", table_data, start_cell="A1")

        # 테이블 목록 조회
        tables = engine.list_tables(test_workbook)

        assert len(tables) > 0
        assert any(t.name == "ProductTable" for t in tables)

    def test_read_table(self, engine, test_workbook):
        """테이블 읽기"""
        # 테이블 생성
        table_data = [["Name", "Score"], ["Alice", 90], ["Bob", 85], ["Charlie", 95]]
        engine.write_table(test_workbook, "Sheet1", "ScoreTable", table_data)

        # 테이블 읽기
        result = engine.read_table(test_workbook, "ScoreTable")

        assert result["table_name"] == "ScoreTable"
        assert len(result["headers"]) == 2
        assert result["row_count"] == 3

    def test_read_table_with_limit(self, engine, test_workbook):
        """테이블 제한 읽기"""
        # 테이블 생성
        table_data = [["ID"]] + [[i] for i in range(1, 11)]  # 10 rows
        engine.write_table(test_workbook, "Sheet1", "LimitTable", table_data)

        # 제한 읽기 (처음 5행)
        result = engine.read_table(test_workbook, "LimitTable", limit=5)

        assert result["row_count"] == 5

    def test_analyze_table(self, engine, test_workbook):
        """테이블 분석"""
        # 테이블 생성
        table_data = [["Item", "Qty"], ["A", 10], ["B", 20]]
        engine.write_table(test_workbook, "Sheet1", "AnalyzeTable", table_data)

        # 분석
        analysis = engine.analyze_table(test_workbook, "AnalyzeTable")

        assert "table_name" in analysis
        assert "row_count" in analysis
        assert "column_count" in analysis


class TestChartOperations:
    """차트 메서드 테스트"""

    def test_add_and_list_charts(self, engine, test_workbook):
        """차트 생성 및 목록 조회"""
        # 차트 데이터 준비
        chart_data = [["Month", "Sales"], ["Jan", 100], ["Feb", 150], ["Mar", 120]]
        engine.write_range(test_workbook, "Sheet1", "A1", chart_data)

        # 차트 생성
        chart_name = engine.add_chart(
            workbook=test_workbook,
            sheet="Sheet1",
            data_range="A1:B4",
            chart_type="column",
            position="D2",
            title="Monthly Sales",
        )

        # 차트 목록 조회
        charts = engine.list_charts(test_workbook)

        assert len(charts) > 0
        assert any(c.name == chart_name for c in charts)

    def test_configure_chart(self, engine, test_workbook):
        """차트 설정"""
        # 차트 생성
        engine.write_range(test_workbook, "Sheet1", "A1", [["X", "Y"], [1, 10], [2, 20]])
        chart_name = engine.add_chart(
            workbook=test_workbook, sheet="Sheet1", data_range="A1:B3", chart_type="line", position="D2"
        )

        # 차트 설정 변경
        engine.configure_chart(workbook=test_workbook, chart_name=chart_name, title="Updated Title", legend_position="bottom")

        # 검증
        charts = engine.list_charts(test_workbook)
        updated_chart = next(c for c in charts if c.name == chart_name)

        assert updated_chart.title == "Updated Title"

    def test_position_chart(self, engine, test_workbook):
        """차트 위치 조정"""
        # 차트 생성
        engine.write_range(test_workbook, "Sheet1", "A1", [[1], [2], [3]])
        chart_name = engine.add_chart(
            workbook=test_workbook, sheet="Sheet1", data_range="A1:A3", chart_type="pie", position="D2"
        )

        # 위치 조정
        new_left, new_top = 200, 100
        new_width, new_height = 500, 400

        engine.position_chart(
            workbook=test_workbook,
            sheet="Sheet1",
            chart_name=chart_name,
            left=new_left,
            top=new_top,
            width=new_width,
            height=new_height,
        )

        # 검증
        charts = engine.list_charts(test_workbook, sheet="Sheet1")
        positioned_chart = next(c for c in charts if c.name == chart_name)

        assert positioned_chart.left == new_left
        assert positioned_chart.width == new_width

    def test_delete_chart(self, engine, test_workbook):
        """차트 삭제"""
        # 차트 생성
        engine.write_range(test_workbook, "Sheet1", "A1", [[1], [2]])
        chart_name = engine.add_chart(
            workbook=test_workbook, sheet="Sheet1", data_range="A1:A2", chart_type="bar", position="D2"
        )

        # 차트 삭제
        engine.delete_chart(test_workbook, "Sheet1", chart_name)

        # 검증
        charts = engine.list_charts(test_workbook, sheet="Sheet1")
        assert not any(c.name == chart_name for c in charts)

    def test_export_chart(self, engine, test_workbook):
        """차트 이미지 내보내기"""
        # 차트 생성
        engine.write_range(test_workbook, "Sheet1", "A1", [["Category", "Value"], ["A", 10], ["B", 20]])
        chart_name = engine.add_chart(
            workbook=test_workbook, sheet="Sheet1", data_range="A1:B3", chart_type="column", position="D2"
        )

        # 이미지로 내보내기
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
            engine.export_chart(test_workbook, "Sheet1", chart_name, tmp.name)

            assert os.path.exists(tmp.name)
            assert os.path.getsize(tmp.name) > 0

            # 정리
            os.unlink(tmp.name)


class TestErrorHandling:
    """에러 처리 테스트"""

    def test_sheet_not_found(self, engine, test_workbook):
        """존재하지 않는 시트 오류"""
        from pyhub_office_automation.excel.engines import SheetNotFoundError

        with pytest.raises(SheetNotFoundError):
            engine.activate_sheet(test_workbook, "NonExistentSheet")

    def test_table_not_found(self, engine, test_workbook):
        """존재하지 않는 테이블 오류"""
        from pyhub_office_automation.excel.engines import TableNotFoundError

        with pytest.raises(TableNotFoundError):
            engine.read_table(test_workbook, "NonExistentTable")

    def test_chart_not_found(self, engine, test_workbook):
        """존재하지 않는 차트 오류"""
        from pyhub_office_automation.excel.engines import ChartNotFoundError

        with pytest.raises(ChartNotFoundError):
            engine.configure_chart(test_workbook, "NonExistentChart")


if __name__ == "__main__":
    pytest.main([__file__, "-v", "--tb=short"])
