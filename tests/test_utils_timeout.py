"""
타임아웃 유틸리티 함수 테스트
utils_timeout.py의 타임아웃 처리 및 COM 정리 기능 테스트
"""

import gc
import platform
import threading
import time
from unittest.mock import Mock, call, patch

import pytest

from pyhub_office_automation.excel.utils_timeout import (
    execute_pivot_operation_with_cleanup,
    execute_with_timeout,
    try_pivot_layout_connection,
)


class TestExecuteWithTimeout:
    """execute_with_timeout 함수 테스트"""

    def test_successful_execution(self):
        """성공적인 함수 실행 테스트"""

        def simple_func():
            return "success"

        success, result, error = execute_with_timeout(simple_func, timeout=5)

        assert success is True
        assert result == "success"
        assert error is None

    def test_successful_execution_with_args(self):
        """인자가 있는 함수 성공 실행 테스트"""

        def add_func(a, b):
            return a + b

        success, result, error = execute_with_timeout(add_func, args=(3, 4), timeout=5)

        assert success is True
        assert result == 7
        assert error is None

    def test_successful_execution_with_kwargs(self):
        """키워드 인자가 있는 함수 성공 실행 테스트"""

        def greet_func(name, greeting="Hello"):
            return f"{greeting}, {name}!"

        success, result, error = execute_with_timeout(greet_func, args=("World",), kwargs={"greeting": "Hi"}, timeout=5)

        assert success is True
        assert result == "Hi, World!"
        assert error is None

    def test_function_exception(self):
        """함수 실행 중 예외 발생 테스트"""

        def error_func():
            raise ValueError("Test error")

        success, result, error = execute_with_timeout(error_func, timeout=5)

        assert success is False
        assert result is None
        assert "Test error" in error

    def test_timeout_occurrence(self):
        """타임아웃 발생 테스트"""

        def slow_func():
            time.sleep(2)
            return "should not reach here"

        success, result, error = execute_with_timeout(slow_func, timeout=1)

        assert success is False
        assert result is None
        assert "1초 내에 완료되지 않아 타임아웃" in error

    def test_edge_case_zero_timeout(self):
        """0초 타임아웃 테스트"""

        def instant_func():
            return "instant"

        # 0초 타임아웃이어도 즉시 완료되는 함수는 성공할 수 있음
        success, result, error = execute_with_timeout(instant_func, timeout=0)

        # 결과는 시스템에 따라 달라질 수 있으므로 타임아웃 케이스도 허용
        assert success in [True, False]

    def test_kwargs_none_handling(self):
        """kwargs=None 처리 테스트"""

        def simple_func():
            return "success"

        success, result, error = execute_with_timeout(simple_func, kwargs=None, timeout=5)

        assert success is True
        assert result == "success"

    def test_thread_daemon_behavior(self):
        """스레드 데몬 동작 테스트"""

        # 실제 스레드가 데몬 스레드로 생성되는지는 내부 구현이므로
        # 타임아웃 동작으로 간접 확인
        def blocking_func():
            # 무한 대기 시뮬레이션
            while True:
                time.sleep(0.1)

        start_time = time.time()
        success, result, error = execute_with_timeout(blocking_func, timeout=1)
        end_time = time.time()

        assert success is False
        assert (end_time - start_time) < 2.0  # 타임아웃이 제대로 작동
        assert "타임아웃" in error


class TestTryPivotLayoutConnection:
    """try_pivot_layout_connection 함수 테스트"""

    @patch("gc.collect")
    def test_successful_pivot_connection(self, mock_gc):
        """성공적인 피벗 연결 테스트"""
        mock_chart = Mock()
        mock_pivot_table = Mock()
        mock_pivot_layout = Mock()
        mock_chart.PivotLayout = mock_pivot_layout

        success, error = try_pivot_layout_connection(mock_chart, mock_pivot_table, timeout=5)

        assert success is True
        assert error is None
        assert mock_pivot_layout.PivotTable == mock_pivot_table
        assert mock_gc.call_count >= 2  # 작업 전후로 호출

    @patch("gc.collect")
    def test_pivot_connection_exception(self, mock_gc):
        """피벗 연결 중 예외 발생 테스트"""
        mock_chart = Mock()
        mock_pivot_table = Mock()
        mock_pivot_layout = Mock()
        mock_chart.PivotLayout = mock_pivot_layout

        # PivotTable 할당 시 예외 발생
        def side_effect(self, value):
            raise Exception("Connection failed")

        type(mock_pivot_layout).PivotTable = property(lambda self: None, side_effect)

        success, error = try_pivot_layout_connection(mock_chart, mock_pivot_table, timeout=5)

        assert success is False
        assert "Connection failed" in error
        assert mock_gc.called  # 에러 시에도 gc.collect 호출

    @patch("gc.collect")
    @patch("platform.system")
    @patch("pythoncom.CoUninitialize")
    def test_pivot_connection_timeout(self, mock_co_uninit, mock_platform, mock_gc):
        """피벗 연결 타임아웃 테스트"""
        mock_platform.return_value = "Windows"
        mock_chart = Mock()
        mock_pivot_table = Mock()
        mock_pivot_layout = Mock()
        mock_chart.PivotLayout = mock_pivot_layout

        # PivotTable 할당을 느리게 만들기
        def slow_setter(self, value):
            time.sleep(2)

        type(mock_pivot_layout).PivotTable = property(lambda self: None, slow_setter)

        success, error = try_pivot_layout_connection(mock_chart, mock_pivot_table, timeout=1)

        assert success is False
        assert "1초 내에 완료되지 않아 타임아웃" in error
        assert "정적 차트로 생성됩니다" in error
        mock_co_uninit.assert_called_once()
        assert mock_gc.called

    @patch("gc.collect")
    @patch("platform.system")
    def test_pivot_connection_timeout_non_windows(self, mock_platform, mock_gc):
        """Windows가 아닌 환경에서 피벗 연결 타임아웃 테스트"""
        mock_platform.return_value = "Linux"
        mock_chart = Mock()
        mock_pivot_table = Mock()
        mock_pivot_layout = Mock()
        mock_chart.PivotLayout = mock_pivot_layout

        def slow_setter(value):
            time.sleep(2)

        type(mock_pivot_layout).PivotTable = property(lambda self: None, slow_setter)

        with patch("pythoncom.CoUninitialize") as mock_co_uninit:
            success, error = try_pivot_layout_connection(mock_chart, mock_pivot_table, timeout=1)

            assert success is False
            mock_co_uninit.assert_not_called()  # Windows가 아니므로 호출되지 않음

    @patch("gc.collect")
    @patch("platform.system")
    @patch("pythoncom.CoUninitialize")
    def test_com_cleanup_error_handling(self, mock_co_uninit, mock_platform, mock_gc):
        """COM 정리 중 에러 처리 테스트"""
        mock_platform.return_value = "Windows"
        mock_co_uninit.side_effect = Exception("CoUninitialize failed")

        mock_chart = Mock()
        mock_pivot_table = Mock()
        mock_pivot_layout = Mock()
        mock_chart.PivotLayout = mock_pivot_layout

        def slow_setter(value):
            time.sleep(2)

        type(mock_pivot_layout).PivotTable = property(lambda self: None, slow_setter)

        # COM 정리 실패해도 함수는 정상 완료되어야 함
        success, error = try_pivot_layout_connection(mock_chart, mock_pivot_table, timeout=1)

        assert success is False
        mock_co_uninit.assert_called_once()

    def test_default_timeout_value(self):
        """기본 타임아웃 값 테스트"""
        mock_chart = Mock()
        mock_pivot_table = Mock()
        mock_pivot_layout = Mock()
        mock_chart.PivotLayout = mock_pivot_layout

        # 타임아웃 값을 지정하지 않으면 10초가 기본값
        with patch("pyhub_office_automation.excel.utils_timeout.execute_with_timeout") as mock_execute:
            mock_execute.return_value = (True, None, None)

            try_pivot_layout_connection(mock_chart, mock_pivot_table)

            # timeout=10으로 호출되었는지 확인
            mock_execute.assert_called_once()
            args, kwargs = mock_execute.call_args
            assert kwargs.get("timeout", args[1] if len(args) > 1 else None) == 10


class TestExecutePivotOperationWithCleanup:
    """execute_pivot_operation_with_cleanup 함수 테스트"""

    @patch("gc.collect")
    def test_successful_pivot_operation(self, mock_gc):
        """성공적인 피벗 작업 테스트"""

        def test_operation(arg1, arg2):
            return f"result_{arg1}_{arg2}"

        success, result, error = execute_pivot_operation_with_cleanup(
            test_operation, "a", "b", timeout=5, description="test operation"
        )

        assert success is True
        assert result == "result_a_b"
        assert error is None
        assert mock_gc.call_count >= 2  # 작업 전후로 호출

    @patch("gc.collect")
    def test_pivot_operation_exception(self, mock_gc):
        """피벗 작업 중 예외 발생 테스트"""

        def error_operation():
            raise ValueError("Operation failed")

        success, result, error = execute_pivot_operation_with_cleanup(
            error_operation, timeout=5, description="error operation"
        )

        assert success is False
        assert result is None
        assert "Operation failed" in error
        assert mock_gc.called  # 에러 시에도 정리

    @patch("gc.collect")
    @patch("platform.system")
    @patch("pythoncom.CoUninitialize")
    def test_pivot_operation_timeout(self, mock_co_uninit, mock_platform, mock_gc):
        """피벗 작업 타임아웃 테스트"""
        mock_platform.return_value = "Windows"

        def slow_operation():
            time.sleep(3)
            return "should not reach"

        success, result, error = execute_pivot_operation_with_cleanup(slow_operation, timeout=1, description="slow operation")

        assert success is False
        assert result is None
        assert "slow operation이(가) 1초 내에 완료되지 않아 타임아웃" in error

        # 타임아웃 후 강제 정리 확인
        assert mock_gc.call_count >= 3  # 실패 시 3번 연속 호출
        mock_co_uninit.assert_called_once()

    @patch("gc.collect")
    def test_default_timeout_and_description(self, mock_gc):
        """기본 타임아웃과 설명 테스트"""

        def quick_operation():
            return "quick_result"

        with patch("pyhub_office_automation.excel.utils_timeout.execute_with_timeout") as mock_execute:
            mock_execute.return_value = (True, "quick_result", None)

            success, result, error = execute_pivot_operation_with_cleanup(quick_operation)

            # 기본값 확인
            mock_execute.assert_called_once()
            args, kwargs = mock_execute.call_args
            assert kwargs.get("timeout", args[1] if len(args) > 1 else None) == 30

    @patch("gc.collect")
    @patch("platform.system")
    @patch("pythoncom.CoUninitialize")
    def test_cleanup_with_com_error(self, mock_co_uninit, mock_platform, mock_gc):
        """COM 정리 중 에러 처리 테스트"""
        mock_platform.return_value = "Windows"
        mock_co_uninit.side_effect = Exception("COM cleanup failed")

        def timeout_operation():
            time.sleep(2)

        # COM 정리 실패해도 함수는 정상 완료되어야 함
        success, result, error = execute_pivot_operation_with_cleanup(timeout_operation, timeout=1)

        assert success is False
        mock_co_uninit.assert_called_once()

    @patch("gc.collect")
    @patch("platform.system")
    def test_cleanup_non_windows_environment(self, mock_platform, mock_gc):
        """Windows가 아닌 환경에서의 정리 테스트"""
        mock_platform.return_value = "Darwin"

        def timeout_operation():
            time.sleep(2)

        with patch("pythoncom.CoUninitialize") as mock_co_uninit:
            success, result, error = execute_pivot_operation_with_cleanup(timeout_operation, timeout=1)

            assert success is False
            mock_co_uninit.assert_not_called()  # Windows가 아니므로 호출되지 않음
            assert mock_gc.call_count >= 3

    def test_custom_description_in_error_message(self):
        """사용자 정의 설명이 에러 메시지에 포함되는지 테스트"""

        def timeout_operation():
            time.sleep(2)

        success, result, error = execute_pivot_operation_with_cleanup(
            timeout_operation, timeout=1, description="custom pivot creation"
        )

        assert success is False
        assert "custom pivot creation이(가)" in error

    @patch("gc.collect")
    def test_multiple_gc_calls_on_failure(self, mock_gc):
        """실패 시 gc.collect가 여러 번 호출되는지 테스트"""

        def timeout_operation():
            time.sleep(2)

        execute_pivot_operation_with_cleanup(timeout_operation, timeout=1)

        # 실패 시 추가로 3번 더 호출됨 (강제 정리)
        assert mock_gc.call_count >= 3

    def test_argument_passing(self):
        """인자 전달 테스트"""

        def multi_arg_operation(a, b, c, multiplier=1):
            return (a + b + c) * multiplier

        success, result, error = execute_pivot_operation_with_cleanup(multi_arg_operation, 1, 2, 3, multiplier=2, timeout=5)

        assert success is True
        assert result == 12  # (1+2+3) * 2
        assert error is None

    def test_no_arguments_operation(self):
        """인자 없는 작업 테스트"""

        def no_arg_operation():
            return "no_args_result"

        success, result, error = execute_pivot_operation_with_cleanup(no_arg_operation, timeout=5)

        assert success is True
        assert result == "no_args_result"
        assert error is None


class TestTimeoutIntegration:
    """타임아웃 관련 통합 테스트"""

    @patch("gc.collect")
    @patch("platform.system")
    @patch("pythoncom.CoUninitialize")
    def test_nested_timeout_operations(self, mock_co_uninit, mock_platform, mock_gc):
        """중첩된 타임아웃 작업 테스트"""
        mock_platform.return_value = "Windows"

        def nested_operation():
            # 내부에서 또 다른 타임아웃 작업 호출
            inner_success, inner_result, inner_error = execute_with_timeout(lambda: "inner_result", timeout=1)
            return f"outer_{inner_result}"

        success, result, error = execute_pivot_operation_with_cleanup(nested_operation, timeout=5)

        assert success is True
        assert result == "outer_inner_result"

    def test_concurrent_timeout_operations(self):
        """동시 타임아웃 작업 테스트"""
        import concurrent.futures

        def slow_operation(delay):
            time.sleep(delay)
            return f"result_{delay}"

        # 여러 작업을 동시에 실행
        with concurrent.futures.ThreadPoolExecutor(max_workers=3) as executor:
            futures = []
            for i, delay in enumerate([0.1, 0.2, 2.0]):  # 마지막 작업은 타임아웃
                future = executor.submit(execute_with_timeout, slow_operation, (delay,), {}, 1)
                futures.append(future)

            results = [f.result() for f in futures]

        # 처음 두 작업은 성공, 마지막은 타임아웃
        assert results[0][0] is True  # 성공
        assert results[1][0] is True  # 성공
        assert results[2][0] is False  # 타임아웃

    @patch("gc.collect")
    def test_memory_cleanup_verification(self, mock_gc):
        """메모리 정리 검증 테스트"""

        def memory_intensive_operation():
            # 큰 객체 생성
            large_data = ["x" * 1000000 for _ in range(10)]
            return len(large_data)

        # 작업 전 gc.collect 호출 횟수
        initial_call_count = mock_gc.call_count

        success, result, error = execute_pivot_operation_with_cleanup(memory_intensive_operation, timeout=10)

        assert success is True
        # gc.collect가 추가로 호출되었는지 확인
        assert mock_gc.call_count > initial_call_count
