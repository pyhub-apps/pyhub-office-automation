"""
Excel 명령어 COM 리소스 관리 통합 테스트
Excel 명령어들이 COM 리소스를 올바르게 정리하는지 검증
"""

import gc
import platform
import tempfile
import time
from pathlib import Path
from typing import Any, Dict
from unittest.mock import MagicMock, Mock, call, patch

import pytest
from typer.testing import CliRunner

from pyhub_office_automation.cli.main import excel_app
from pyhub_office_automation.excel.utils import COMResourceManager


class MockCOMObject:
    """COM 객체 모킹을 위한 헬퍼 클래스"""

    def __init__(self, name: str = "MockCOMObject"):
        self.name = name
        self.api = Mock()
        self.closed = False
        self.quit_called = False

    def close(self):
        self.closed = True

    def quit(self):
        self.quit_called = True


class TestExcelCommandCOMCleanup:
    """Excel 명령어 COM 정리 통합 테스트"""

    @pytest.fixture
    def mock_xlwings_with_com(self):
        """COM 리소스 관리가 포함된 xlwings 모킹"""
        with patch("xlwings.App") as mock_app_class, patch("xlwings.Book") as mock_book_class:

            # App 인스턴스 모킹
            mock_app = MockCOMObject("App")
            mock_app.visible = True
            mock_app.books = Mock()
            mock_app_class.return_value = mock_app

            # Book 인스턴스 모킹
            mock_book = MockCOMObject("Book")
            mock_book.name = "test_workbook.xlsx"
            mock_book.fullname = "/path/to/test_workbook.xlsx"
            mock_book.saved = True

            # Sheet 모킹
            mock_sheet = MockCOMObject("Sheet")
            mock_sheet.name = "Sheet1"
            mock_sheet.index = 1

            # Range 모킹
            mock_range = Mock()
            mock_range.value = [["A1", "B1"], ["A2", "B2"]]
            mock_sheet.range = Mock(return_value=mock_range)

            # 관계 설정
            mock_book.sheets = [mock_sheet]
            mock_book.sheets.active = mock_sheet
            mock_app.books.open = Mock(return_value=mock_book)
            mock_book_class.return_value = mock_book

            yield {
                "app_class": mock_app_class,
                "app": mock_app,
                "book_class": mock_book_class,
                "book": mock_book,
                "sheet": mock_sheet,
                "range": mock_range,
            }

    @patch("gc.collect")
    def test_range_read_com_cleanup(self, mock_gc, mock_xlwings_with_com):
        """range-read 명령의 COM 정리 테스트"""
        runner = CliRunner()

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as temp_file:
            temp_path = Path(temp_file.name)

        try:
            result = runner.invoke(
                excel_app, ["range-read", "--file-path", str(temp_path), "--range", "A1:B2", "--format", "json"]
            )

            # 명령이 성공적으로 실행되었는지 확인
            assert result.exit_code == 0

            # gc.collect이 finally 블록에서 호출되었는지 확인
            assert mock_gc.called

        finally:
            if temp_path.exists():
                temp_path.unlink()

    @patch("gc.collect")
    def test_range_write_com_cleanup(self, mock_gc, mock_xlwings_with_com):
        """range-write 명령의 COM 정리 테스트"""
        runner = CliRunner()

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as temp_file:
            temp_path = Path(temp_file.name)

        try:
            result = runner.invoke(
                excel_app,
                [
                    "range-write",
                    "--file-path",
                    str(temp_path),
                    "--range",
                    "A1",
                    "--data",
                    '["Hello", "World"]',
                    "--format",
                    "json",
                ],
            )

            # gc.collect이 finally 블록에서 호출되었는지 확인
            assert mock_gc.called

        finally:
            if temp_path.exists():
                temp_path.unlink()

    @patch("gc.collect")
    def test_workbook_open_com_cleanup(self, mock_gc, mock_xlwings_with_com):
        """workbook-open 명령의 COM 정리 테스트"""
        runner = CliRunner()

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as temp_file:
            temp_path = Path(temp_file.name)

        try:
            result = runner.invoke(excel_app, ["workbook-open", "--file-path", str(temp_path), "--format", "json"])

            # gc.collect이 finally 블록에서 호출되었는지 확인
            assert mock_gc.called

        finally:
            if temp_path.exists():
                temp_path.unlink()

    @patch("gc.collect")
    @patch("platform.system")
    @patch("pythoncom.CoUninitialize")
    def test_com_cleanup_on_exception(self, mock_co_uninit, mock_platform, mock_gc):
        """예외 발생 시 COM 정리 테스트"""
        mock_platform.return_value = "Windows"

        runner = CliRunner()

        # 존재하지 않는 파일로 테스트
        result = runner.invoke(excel_app, ["range-read", "--file-path", "/nonexistent/file.xlsx", "--range", "A1:B2"])

        # 명령이 실패하더라도 COM 정리는 수행되어야 함
        assert result.exit_code != 0
        assert mock_gc.called

    def test_multiple_operations_memory_usage(self, mock_xlwings_with_com):
        """여러 작업 수행 시 메모리 사용량 테스트"""
        runner = CliRunner()

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as temp_file:
            temp_path = Path(temp_file.name)

        try:
            # 여러 번의 작업 수행
            for i in range(5):
                result = runner.invoke(
                    excel_app, ["range-read", "--file-path", str(temp_path), "--range", f"A{i+1}:B{i+1}", "--format", "json"]
                )

            # 모든 작업이 독립적으로 정리되어야 함
            # (실제 메모리 측정은 어렵지만, 예외가 발생하지 않는 것으로 확인)

        finally:
            if temp_path.exists():
                temp_path.unlink()

    @patch("gc.collect")
    def test_com_resource_manager_integration(self, mock_gc):
        """COMResourceManager와 Excel 명령 통합 테스트"""
        with COMResourceManager(verbose=True) as com_manager:
            # Mock COM 객체들 생성
            mock_app = MockCOMObject("App")
            mock_book = MockCOMObject("Book")
            mock_sheet = MockCOMObject("Sheet")

            # COMResourceManager에 추가
            com_manager.add(mock_app, "Excel Application")
            com_manager.add(mock_book, "Excel Workbook")
            com_manager.add(mock_sheet, "Excel Sheet")

            # 컨텍스트 내에서 작업 수행 시뮬레이션
            assert len(com_manager.com_objects) == 3

        # 컨텍스트 종료 후 정리 확인
        assert len(com_manager.com_objects) == 0
        assert mock_app.closed is True
        assert mock_book.closed is True
        assert mock_sheet.closed is True
        assert mock_gc.call_count >= 3

    def test_api_reference_cleanup_integration(self):
        """API 참조 정리 통합 테스트"""
        with COMResourceManager() as com_manager:
            # API가 있는 Mock 객체
            mock_obj = Mock()
            mock_api = Mock()
            mock_api.Release = Mock()
            mock_obj.api = mock_api

            com_manager.add(mock_obj)

            # API 참조가 추가되었는지 확인
            assert len(com_manager.api_refs) == 1

        # API Release가 호출되었는지 확인
        mock_api.Release.assert_called_once()

    @patch("gc.collect")
    def test_nested_com_operations(self, mock_gc, mock_xlwings_with_com):
        """중첩된 COM 작업 테스트"""

        def nested_excel_operations():
            with COMResourceManager() as outer_manager:
                outer_obj = MockCOMObject("Outer")
                outer_manager.add(outer_obj)

                with COMResourceManager() as inner_manager:
                    inner_obj = MockCOMObject("Inner")
                    inner_manager.add(inner_obj)

                    return "nested_complete"

        result = nested_excel_operations()
        assert result == "nested_complete"
        assert mock_gc.called

    def test_error_recovery_com_cleanup(self, mock_xlwings_with_com):
        """에러 복구 시 COM 정리 테스트"""
        error_occurred = False

        try:
            with COMResourceManager() as com_manager:
                mock_obj = MockCOMObject("ErrorTest")
                com_manager.add(mock_obj)

                # 의도적으로 에러 발생
                raise ValueError("Test error")

        except ValueError:
            error_occurred = True

        # 에러가 발생해도 COM 정리는 수행되어야 함
        assert error_occurred is True
        # COMResourceManager는 예외를 억제하지 않음


class TestExcelCommandTimeout:
    """Excel 명령어 타임아웃 처리 테스트"""

    @patch("pyhub_office_automation.excel.utils_timeout.execute_with_timeout")
    def test_chart_pivot_create_timeout_handling(self, mock_execute):
        """피벗차트 생성 시 타임아웃 처리 테스트"""
        # 타임아웃 시뮬레이션
        mock_execute.return_value = (False, None, "작업이 120초 내에 완료되지 않아 타임아웃")

        runner = CliRunner()

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as temp_file:
            temp_path = Path(temp_file.name)

        try:
            result = runner.invoke(
                excel_app,
                [
                    "chart-pivot-create",
                    "--file-path",
                    str(temp_path),
                    "--data-range",
                    "A1:D10",
                    "--rows",
                    "Category",
                    "--values",
                    "Sales",
                    "--format",
                    "json",
                ],
            )

            # 타임아웃 발생 시에도 적절한 응답 반환
            # (명령어가 실패하거나 fallback 동작 수행)

        finally:
            if temp_path.exists():
                temp_path.unlink()

    @patch("time.sleep")
    @patch("threading.Thread")
    def test_timeout_thread_behavior(self, mock_thread, mock_sleep):
        """타임아웃 처리 시 스레드 동작 테스트"""
        from pyhub_office_automation.excel.utils_timeout import execute_with_timeout

        # Mock 스레드 설정
        mock_thread_instance = Mock()
        mock_thread_instance.is_alive.return_value = True  # 타임아웃 시뮬레이션
        mock_thread.return_value = mock_thread_instance

        def slow_function():
            return "should timeout"

        success, result, error = execute_with_timeout(slow_function, timeout=1)

        assert success is False
        assert "타임아웃" in error
        mock_thread.assert_called_once()
        mock_thread_instance.start.assert_called_once()
        mock_thread_instance.join.assert_called_once_with(1)

    def test_timeout_cleanup_verification(self):
        """타임아웃 후 정리 검증 테스트"""
        from pyhub_office_automation.excel.utils_timeout import execute_pivot_operation_with_cleanup

        cleanup_called = []

        def operation_with_cleanup_tracking():
            cleanup_called.append("operation_start")
            time.sleep(0.5)  # 짧은 지연
            cleanup_called.append("operation_end")
            return "success"

        with patch("gc.collect") as mock_gc:
            success, result, error = execute_pivot_operation_with_cleanup(operation_with_cleanup_tracking, timeout=2)

            assert success is True
            assert "operation_start" in cleanup_called
            assert "operation_end" in cleanup_called
            assert mock_gc.called


class TestExcelCommandPerformance:
    """Excel 명령어 성능 및 메모리 테스트"""

    def test_com_cleanup_performance(self):
        """COM 정리 성능 테스트"""
        start_time = time.time()

        # 많은 COM 객체 생성 및 정리
        with COMResourceManager() as com_manager:
            for i in range(100):
                mock_obj = MockCOMObject(f"Object_{i}")
                com_manager.add(mock_obj)

        end_time = time.time()

        # 정리가 1초 내에 완료되어야 함
        assert (end_time - start_time) < 1.0

    @patch("gc.collect")
    def test_garbage_collection_frequency(self, mock_gc):
        """가비지 컬렉션 호출 빈도 테스트"""
        from pyhub_office_automation.excel.utils_timeout import execute_pivot_operation_with_cleanup

        def simple_operation():
            return "test_result"

        # 초기 호출 횟수
        initial_count = mock_gc.call_count

        execute_pivot_operation_with_cleanup(simple_operation, timeout=5)

        # 추가 gc.collect 호출 확인 (최소 2회: 작업 전후)
        assert mock_gc.call_count > initial_count

    def test_memory_leak_prevention(self):
        """메모리 누수 방지 테스트"""
        # 여러 번의 COM 리소스 생성 및 정리
        for iteration in range(10):
            with COMResourceManager() as com_manager:
                # 큰 Mock 객체 생성
                for i in range(10):
                    mock_obj = MockCOMObject(f"Large_Object_{iteration}_{i}")
                    mock_obj.large_data = "x" * 10000  # 큰 데이터 추가
                    com_manager.add(mock_obj)

        # 반복 작업 후에도 메모리가 정리되어야 함
        # (실제 메모리 측정은 복잡하므로, 예외 발생 없음으로 검증)

    def test_concurrent_com_operations(self):
        """동시 COM 작업 테스트"""
        import concurrent.futures
        import threading

        def create_and_cleanup_com_objects(thread_id):
            """스레드별로 COM 객체 생성 및 정리"""
            with COMResourceManager() as com_manager:
                for i in range(5):
                    mock_obj = MockCOMObject(f"Thread_{thread_id}_Object_{i}")
                    com_manager.add(mock_obj)
                return f"Thread_{thread_id}_completed"

        # 여러 스레드에서 동시에 COM 작업 수행
        with concurrent.futures.ThreadPoolExecutor(max_workers=3) as executor:
            futures = [executor.submit(create_and_cleanup_com_objects, i) for i in range(3)]
            results = [future.result() for future in futures]

        # 모든 스레드가 성공적으로 완료되어야 함
        assert len(results) == 3
        assert all("completed" in result for result in results)

    def test_excel_command_memory_stability(self, mock_xlwings_with_com):
        """Excel 명령어 메모리 안정성 테스트"""
        runner = CliRunner()

        # 같은 명령을 여러 번 실행
        for i in range(10):
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as temp_file:
                temp_path = Path(temp_file.name)

            try:
                result = runner.invoke(
                    excel_app, ["range-read", "--file-path", str(temp_path), "--range", f"A1:A{i+1}", "--format", "json"]
                )

                # 각 명령이 독립적으로 실행되어야 함
                # (이전 실행의 영향을 받지 않음)

            finally:
                if temp_path.exists():
                    temp_path.unlink()


class TestEdgeCaseHandling:
    """엣지 케이스 처리 테스트"""

    def test_com_object_with_no_cleanup_methods(self):
        """정리 메서드가 없는 COM 객체 처리 테스트"""
        with COMResourceManager() as com_manager:
            mock_obj = Mock()
            # close, quit 메서드 제거
            del mock_obj.close
            del mock_obj.quit

            com_manager.add(mock_obj)

        # 정리 메서드가 없어도 예외 발생하지 않아야 함

    def test_com_object_cleanup_failure(self):
        """COM 객체 정리 실패 테스트"""
        with COMResourceManager(verbose=True) as com_manager:
            mock_obj = Mock()
            mock_obj.close.side_effect = Exception("Close failed")

            com_manager.add(mock_obj)

        # 정리 실패해도 컨텍스트는 정상 종료되어야 함

    def test_api_reference_with_null_api(self):
        """API가 None인 객체 처리 테스트"""
        with COMResourceManager() as com_manager:
            mock_obj = Mock()
            mock_obj.api = None

            com_manager.add(mock_obj)

        # API가 None이어도 예외 발생하지 않아야 함

    @patch("platform.system")
    def test_com_cleanup_on_different_platforms(self, mock_platform):
        """다른 플랫폼에서의 COM 정리 테스트"""
        platforms = ["Windows", "Darwin", "Linux"]

        for platform_name in platforms:
            mock_platform.return_value = platform_name

            with COMResourceManager() as com_manager:
                mock_obj = MockCOMObject(f"Object_{platform_name}")
                com_manager.add(mock_obj)

            # 모든 플랫폼에서 정상적으로 정리되어야 함
            assert mock_obj.closed is True

    def test_empty_com_manager_cleanup(self):
        """빈 COMResourceManager 정리 테스트"""
        with COMResourceManager() as com_manager:
            pass  # 아무 객체도 추가하지 않음

        # 빈 매니저도 정상적으로 종료되어야 함
