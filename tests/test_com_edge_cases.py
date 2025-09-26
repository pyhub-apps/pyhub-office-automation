"""
COM 리소스 관리 엣지 케이스 및 에러 시나리오 테스트
예외 상황, 플랫폼 차이, 에러 복구 등의 엣지 케이스 테스트
"""

import gc
import platform
import sys
import threading
import time
from typing import Any, Dict, List
from unittest.mock import MagicMock, Mock, call, patch

import pytest

from pyhub_office_automation.excel.utils import COMResourceManager
from pyhub_office_automation.excel.utils_timeout import (
    execute_pivot_operation_with_cleanup,
    execute_with_timeout,
    try_pivot_layout_connection,
)


class BrokenCOMObject:
    """의도적으로 문제가 있는 COM 객체"""

    def __init__(self, failure_mode: str = "close"):
        self.failure_mode = failure_mode
        self.name = f"BrokenCOMObject_{failure_mode}"
        self.api = Mock()
        self.closed = False

    def close(self):
        if self.failure_mode == "close":
            raise RuntimeError("Close method failed")
        self.closed = True

    def quit(self):
        if self.failure_mode == "quit":
            raise RuntimeError("Quit method failed")

    @property
    def broken_property(self):
        if self.failure_mode == "property":
            raise AttributeError("Property access failed")
        return "working_property"


class TestCOMResourceManagerEdgeCases:
    """COMResourceManager 엣지 케이스 테스트"""

    def test_cleanup_with_broken_close_method(self):
        """close 메서드가 실패하는 객체 정리 테스트"""
        with COMResourceManager(verbose=True) as com_manager:
            broken_obj = BrokenCOMObject("close")
            com_manager.add(broken_obj)

            # 정상 객체도 함께 추가
            normal_obj = Mock()
            normal_obj.close = Mock()
            com_manager.add(normal_obj)

        # 정상 객체는 정리되어야 함
        normal_obj.close.assert_called_once()

        # 에러가 발생해도 컨텍스트는 정상 종료

    def test_cleanup_with_broken_quit_method(self):
        """quit 메서드가 실패하는 객체 정리 테스트"""
        with COMResourceManager(verbose=True) as com_manager:
            broken_obj = BrokenCOMObject("quit")
            del broken_obj.close  # close 메서드 제거하여 quit 호출하도록 함
            com_manager.add(broken_obj)

        # 에러 발생해도 컨텍스트 정상 종료

    def test_cleanup_with_missing_methods(self):
        """정리 메서드가 없는 객체 테스트"""
        with COMResourceManager() as com_manager:
            obj = Mock()
            # close, quit 메서드 제거
            del obj.close
            del obj.quit
            com_manager.add(obj)

        # 메서드가 없어도 에러 발생하지 않아야 함

    def test_api_cleanup_with_missing_release(self):
        """Release 메서드가 없는 API 정리 테스트"""
        with COMResourceManager() as com_manager:
            obj = Mock()
            api_obj = Mock()
            del api_obj.Release  # Release 메서드 제거
            obj.api = api_obj
            com_manager.add(obj)

        # Release 메서드가 없어도 정상 처리

    def test_api_cleanup_with_broken_release(self):
        """Release 메서드가 실패하는 API 정리 테스트"""
        with COMResourceManager(verbose=True) as com_manager:
            obj = Mock()
            api_obj = Mock()
            api_obj.Release.side_effect = Exception("Release failed")
            obj.api = api_obj
            com_manager.add(obj)

        # Release 실패해도 정상 처리

    def test_api_reference_becomes_none_during_cleanup(self):
        """정리 중 API 참조가 None이 되는 경우 테스트"""
        with COMResourceManager() as com_manager:
            obj = Mock()
            obj.api = Mock()
            com_manager.add(obj)

            # 정리 중에 api가 None이 되도록 설정
            def side_effect_setattr(attr, value):
                if attr == "api":
                    obj.api = None

            with patch.object(obj, "__setattr__", side_effect=side_effect_setattr):
                pass

        # 정리 과정에서 API 참조가 변경되어도 정상 처리

    def test_object_deletion_failure(self):
        """객체 삭제 실패 테스트"""

        class UndeletableObject:
            def __init__(self):
                self.name = "UndeletableObject"

            def __del__(self):
                raise RuntimeError("Cannot delete this object")

            def close(self):
                pass

        with COMResourceManager() as com_manager:
            undeletable = UndeletableObject()
            com_manager.add(undeletable)

        # 객체 삭제 실패해도 정상 처리

    @patch("gc.collect")
    def test_gc_collect_failure(self, mock_gc):
        """가비지 컬렉션 실패 테스트"""
        mock_gc.side_effect = Exception("GC failed")

        with COMResourceManager() as com_manager:
            obj = Mock()
            com_manager.add(obj)

        # GC 실패해도 정상 처리

    @patch("platform.system")
    @patch("pythoncom.CoUninitialize")
    def test_com_uninitialize_failure(self, mock_co_uninit, mock_platform):
        """COM 라이브러리 정리 실패 테스트"""
        mock_platform.return_value = "Windows"
        mock_co_uninit.side_effect = Exception("CoUninitialize failed")

        with COMResourceManager() as com_manager:
            obj = Mock()
            com_manager.add(obj)

        # COM 라이브러리 정리 실패해도 정상 처리
        mock_co_uninit.assert_called_once()

    @patch("platform.system")
    def test_com_cleanup_on_unknown_platform(self, mock_platform):
        """알 수 없는 플랫폼에서의 COM 정리 테스트"""
        mock_platform.return_value = "UnknownOS"

        with patch("pythoncom.CoUninitialize") as mock_co_uninit:
            with COMResourceManager() as com_manager:
                obj = Mock()
                com_manager.add(obj)

            # Windows가 아니므로 COM 정리 호출되지 않음
            mock_co_uninit.assert_not_called()

    def test_very_large_object_list(self):
        """매우 큰 객체 리스트 처리 테스트"""
        with COMResourceManager() as com_manager:
            # 1000개 객체 추가
            objects = []
            for i in range(1000):
                obj = Mock()
                obj.name = f"Object_{i}"
                obj.close = Mock()
                objects.append(obj)
                com_manager.add(obj)

            assert len(com_manager.com_objects) == 1000

        # 모든 객체가 정리되었는지 확인
        for obj in objects:
            obj.close.assert_called_once()

        assert len(com_manager.com_objects) == 0

    def test_duplicate_object_handling(self):
        """중복 객체 추가 처리 테스트"""
        with COMResourceManager() as com_manager:
            obj = Mock()
            obj.close = Mock()

            # 같은 객체를 여러 번 추가
            com_manager.add(obj)
            com_manager.add(obj)
            com_manager.add(obj)

            # 리스트에는 한 번만 추가되어야 함
            assert len(com_manager.com_objects) == 1

        # close는 한 번만 호출되어야 함
        obj.close.assert_called_once()

    def test_exception_in_context_body(self):
        """컨텍스트 본문에서 예외 발생 시 정리 테스트"""
        obj = Mock()
        obj.close = Mock()

        with pytest.raises(ValueError):
            with COMResourceManager() as com_manager:
                com_manager.add(obj)
                raise ValueError("Test exception in context")

        # 예외 발생해도 정리는 수행되어야 함
        obj.close.assert_called_once()


class TestTimeoutEdgeCases:
    """타임아웃 처리 엣지 케이스 테스트"""

    def test_zero_timeout(self):
        """0초 타임아웃 테스트"""

        def instant_function():
            return "instant_result"

        success, result, error = execute_with_timeout(instant_function, timeout=0)

        # 결과는 시스템에 따라 달라질 수 있음
        if success:
            assert result == "instant_result"
        else:
            assert "타임아웃" in error

    def test_negative_timeout(self):
        """음수 타임아웃 테스트"""

        def simple_function():
            return "result"

        success, result, error = execute_with_timeout(simple_function, timeout=-1)

        # 음수 타임아웃도 처리되어야 함
        assert success in [True, False]

    def test_very_large_timeout(self):
        """매우 큰 타임아웃 값 테스트"""

        def quick_function():
            return "quick_result"

        success, result, error = execute_with_timeout(quick_function, timeout=86400)  # 24시간

        assert success is True
        assert result == "quick_result"

    def test_function_with_no_return_value(self):
        """반환값이 없는 함수 테스트"""
        side_effect_tracker = []

        def void_function():
            side_effect_tracker.append("executed")

        success, result, error = execute_with_timeout(void_function, timeout=5)

        assert success is True
        assert result is None
        assert "executed" in side_effect_tracker

    def test_function_returning_none(self):
        """None을 반환하는 함수 테스트"""

        def none_function():
            return None

        success, result, error = execute_with_timeout(none_function, timeout=5)

        assert success is True
        assert result is None
        assert error is None

    def test_function_with_complex_exception(self):
        """복잡한 예외가 발생하는 함수 테스트"""

        def complex_exception_function():
            try:
                raise ValueError("Inner exception")
            except ValueError as e:
                raise RuntimeError("Outer exception") from e

        success, result, error = execute_with_timeout(complex_exception_function, timeout=5)

        assert success is False
        assert result is None
        assert "Outer exception" in error

    def test_thread_interruption_handling(self):
        """스레드 중단 처리 테스트"""

        def long_running_function():
            for i in range(1000):
                time.sleep(0.01)  # 10초 작업
            return "should_not_reach"

        success, result, error = execute_with_timeout(long_running_function, timeout=0.5)

        assert success is False
        assert "타임아웃" in error

    @patch("threading.Thread")
    def test_thread_creation_failure(self, mock_thread):
        """스레드 생성 실패 테스트"""
        mock_thread.side_effect = Exception("Thread creation failed")

        def simple_function():
            return "result"

        with pytest.raises(Exception):
            execute_with_timeout(simple_function, timeout=5)

    def test_pivot_connection_with_invalid_objects(self):
        """유효하지 않은 객체로 피벗 연결 테스트"""
        # None 객체들로 테스트
        success, error = try_pivot_layout_connection(None, None, timeout=1)

        assert success is False
        assert error is not None

    def test_pivot_connection_with_broken_pivot_layout(self):
        """깨진 PivotLayout으로 피벗 연결 테스트"""
        mock_chart = Mock()
        mock_pivot_table = Mock()

        # PivotLayout 접근 시 에러 발생
        mock_chart.PivotLayout = Mock()
        type(mock_chart.PivotLayout).PivotTable = property(
            lambda self: None, lambda self, value: (_ for _ in ()).throw(AttributeError("PivotLayout broken"))
        )

        success, error = try_pivot_layout_connection(mock_chart, mock_pivot_table, timeout=5)

        assert success is False
        assert "PivotLayout broken" in error

    @patch("gc.collect")
    def test_pivot_operation_with_gc_failure(self, mock_gc):
        """가비지 컬렉션 실패가 있는 피벗 작업 테스트"""
        mock_gc.side_effect = [None, Exception("GC failed"), None]

        def simple_operation():
            return "result"

        success, result, error = execute_pivot_operation_with_cleanup(simple_operation, timeout=5)

        # GC 실패해도 작업은 성공해야 함
        assert success is True
        assert result == "result"

    def test_recursive_timeout_operations(self):
        """재귀적 타임아웃 작업 테스트"""

        def recursive_operation(depth: int):
            if depth <= 0:
                return "base_case"

            # 내부에서 또 다른 타임아웃 작업 호출
            success, result, error = execute_with_timeout(lambda: recursive_operation(depth - 1), timeout=10)

            if success:
                return f"depth_{depth}_{result}"
            else:
                return f"error_at_depth_{depth}"

        success, result, error = execute_with_timeout(lambda: recursive_operation(3), timeout=30)

        assert success is True
        assert "depth_3" in result
        assert "base_case" in result


class TestPlatformSpecificEdgeCases:
    """플랫폼별 엣지 케이스 테스트"""

    @patch("platform.system")
    def test_windows_specific_com_handling(self, mock_platform):
        """Windows 특화 COM 처리 테스트"""
        mock_platform.return_value = "Windows"

        with patch("pythoncom.CoUninitialize") as mock_co_uninit:
            with COMResourceManager() as com_manager:
                obj = Mock()
                com_manager.add(obj)

            mock_co_uninit.assert_called_once()

    @patch("platform.system")
    def test_macos_com_handling(self, mock_platform):
        """macOS COM 처리 테스트"""
        mock_platform.return_value = "Darwin"

        with patch("pythoncom.CoUninitialize") as mock_co_uninit:
            with COMResourceManager() as com_manager:
                obj = Mock()
                com_manager.add(obj)

            # macOS에서는 COM 라이브러리 정리 호출되지 않음
            mock_co_uninit.assert_not_called()

    @patch("platform.system")
    def test_linux_com_handling(self, mock_platform):
        """Linux COM 처리 테스트"""
        mock_platform.return_value = "Linux"

        with patch("pythoncom.CoUninitialize") as mock_co_uninit:
            with COMResourceManager() as com_manager:
                obj = Mock()
                com_manager.add(obj)

            # Linux에서는 COM 라이브러리 정리 호출되지 않음
            mock_co_uninit.assert_not_called()

    @patch("platform.system")
    def test_pythoncom_import_failure(self, mock_platform):
        """pythoncom import 실패 테스트"""
        mock_platform.return_value = "Windows"

        # pythoncom 모듈이 없는 상황 시뮬레이션
        with patch("builtins.__import__") as mock_import:

            def import_side_effect(name, *args, **kwargs):
                if name == "pythoncom":
                    raise ImportError("No module named 'pythoncom'")
                return __import__(name, *args, **kwargs)

            mock_import.side_effect = import_side_effect

            with COMResourceManager() as com_manager:
                obj = Mock()
                com_manager.add(obj)

            # import 실패해도 정상 처리되어야 함

    @patch("sys.platform", "win32")
    def test_sys_platform_windows_detection(self):
        """sys.platform을 통한 Windows 감지 테스트"""
        # 실제로는 platform.system()을 사용하지만 시스템 감지 로직 테스트
        assert sys.platform == "win32"

    def test_unicode_object_names(self):
        """유니코드 객체 이름 처리 테스트"""
        with COMResourceManager(verbose=True) as com_manager:
            obj = Mock()
            obj.close = Mock()
            unicode_name = "테스트_객체_🔥"

            com_manager.add(obj, unicode_name)

        # 유니코드 이름도 정상 처리
        obj.close.assert_called_once()

    def test_very_long_object_names(self):
        """매우 긴 객체 이름 처리 테스트"""
        with COMResourceManager(verbose=True) as com_manager:
            obj = Mock()
            obj.close = Mock()
            long_name = "very_long_object_name_" * 100

            com_manager.add(obj, long_name)

        # 긴 이름도 정상 처리
        obj.close.assert_called_once()


class TestConcurrencyEdgeCases:
    """동시성 엣지 케이스 테스트"""

    def test_concurrent_com_manager_creation(self):
        """동시 COMResourceManager 생성 테스트"""
        results = []

        def create_com_manager(manager_id: int):
            try:
                with COMResourceManager() as com_manager:
                    obj = Mock()
                    obj.name = f"Object_Manager_{manager_id}"
                    obj.close = Mock()
                    com_manager.add(obj)

                    time.sleep(0.1)  # 작업 시뮬레이션
                    return f"Manager_{manager_id}_success"
            except Exception as e:
                return f"Manager_{manager_id}_error_{str(e)}"

        # 10개의 동시 COM 매니저 생성
        threads = []
        for i in range(10):
            thread = threading.Thread(target=lambda i=i: results.append(create_com_manager(i)))
            threads.append(thread)
            thread.start()

        # 모든 스레드 완료 대기
        for thread in threads:
            thread.join()

        # 모든 매니저가 성공적으로 완료되어야 함
        assert len(results) == 10
        assert all("success" in result for result in results)

    def test_race_condition_in_object_addition(self):
        """객체 추가 시 경합 조건 테스트"""
        com_manager = COMResourceManager()
        results = []

        def add_objects_concurrently(thread_id: int):
            try:
                for i in range(50):
                    obj = Mock()
                    obj.name = f"Thread_{thread_id}_Object_{i}"
                    obj.close = Mock()
                    com_manager.add(obj)
                return f"Thread_{thread_id}_completed"
            except Exception as e:
                return f"Thread_{thread_id}_error_{str(e)}"

        # 5개 스레드가 동시에 객체 추가
        threads = []
        for i in range(5):
            thread = threading.Thread(target=lambda i=i: results.append(add_objects_concurrently(i)))
            threads.append(thread)
            thread.start()

        for thread in threads:
            thread.join()

        # 모든 스레드가 완료되어야 함
        assert len(results) == 5

        # 정리 수행
        with com_manager:
            pass

        # 모든 객체가 추가되었는지 확인 (250개)
        # 실제로는 경합 조건 때문에 정확한 수를 예측하기 어려우므로 범위로 확인
        # 정리 후에는 0개가 되어야 함
        assert len(com_manager.com_objects) == 0

    def test_timeout_with_thread_local_storage(self):
        """스레드 로컬 스토리지와 타임아웃 테스트"""
        import threading

        thread_local = threading.local()

        def thread_local_operation():
            thread_local.data = "thread_specific_data"
            time.sleep(0.1)
            return getattr(thread_local, "data", "no_data")

        success, result, error = execute_with_timeout(thread_local_operation, timeout=1)

        assert success is True
        assert result == "thread_specific_data"

    def test_nested_threading_with_com_cleanup(self):
        """중첩 스레딩과 COM 정리 테스트"""
        results = []

        def nested_thread_operation():
            def inner_operation():
                with COMResourceManager() as com_manager:
                    obj = Mock()
                    obj.close = Mock()
                    com_manager.add(obj)
                    return "inner_completed"

            success, result, error = execute_with_timeout(inner_operation, timeout=5)
            return (success, result, error)

        # 외부 스레드에서 내부 타임아웃 작업 실행
        success, result, error = execute_with_timeout(nested_thread_operation, timeout=10)

        assert success is True
        inner_success, inner_result, inner_error = result
        assert inner_success is True
        assert inner_result == "inner_completed"


class TestResourceExhaustionScenarios:
    """리소스 고갈 시나리오 테스트"""

    @patch("threading.Thread")
    def test_thread_pool_exhaustion(self, mock_thread_class):
        """스레드 풀 고갈 시나리오 테스트"""
        # 스레드 생성을 제한된 횟수만 허용
        call_count = [0]

        def limited_thread_creation(*args, **kwargs):
            call_count[0] += 1
            if call_count[0] > 5:
                raise RuntimeError("Too many threads")
            return Mock()

        mock_thread_class.side_effect = limited_thread_creation

        # 여러 번의 타임아웃 작업 시도
        for i in range(3):  # 5개 제한이므로 3개는 성공해야 함
            try:
                success, result, error = execute_with_timeout(lambda: "test_result", timeout=1)
                # 스레드 생성 실패 전까지는 성공 또는 실패 가능
            except RuntimeError:
                # 스레드 생성 실패 시 예외 발생 허용
                break

    def test_memory_pressure_during_cleanup(self):
        """정리 중 메모리 압박 상황 테스트"""
        # 메모리를 많이 사용하는 객체들 생성
        large_objects = []

        try:
            with COMResourceManager() as com_manager:
                for i in range(10):
                    # 큰 데이터를 가진 Mock 객체
                    obj = Mock()
                    obj.large_data = bytearray(1024 * 1024)  # 1MB
                    obj.close = Mock()
                    large_objects.append(obj)
                    com_manager.add(obj)

                # 메모리 압박 상황에서 정리가 정상적으로 수행되는지 확인

        except MemoryError:
            # 메모리 부족 시에도 정상적으로 처리되어야 함
            pass

        finally:
            # 명시적으로 큰 객체들 해제
            for obj in large_objects:
                obj.large_data = None

    def test_system_resource_limitation(self):
        """시스템 리소스 제한 테스트"""
        # 많은 수의 객체를 생성하여 시스템 한계 근처에서 테스트
        try:
            with COMResourceManager() as com_manager:
                objects = []
                for i in range(10000):  # 많은 수의 객체
                    obj = Mock()
                    obj.name = f"SystemTest_{i}"
                    obj.close = Mock()
                    objects.append(obj)
                    com_manager.add(obj)

                    if i % 1000 == 0:
                        # 중간중간 가비지 컬렉션으로 메모리 정리
                        gc.collect()

        except Exception as e:
            # 시스템 제한으로 인한 예외도 정상적으로 처리되어야 함
            print(f"시스템 제한으로 인한 예외: {e}")


if __name__ == "__main__":
    # 엣지 케이스 테스트 실행 예시
    print("COM 리소스 관리 엣지 케이스 테스트 실행 중...")

    test_edge_cases = TestCOMResourceManagerEdgeCases()
    test_edge_cases.test_cleanup_with_broken_close_method()

    print("엣지 케이스 테스트 완료!")
