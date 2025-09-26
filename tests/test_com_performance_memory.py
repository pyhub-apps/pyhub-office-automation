"""
COM 리소스 관리 성능 및 메모리 테스트
메모리 누수 감지, 성능 측정, 대용량 데이터 처리 테스트
"""

import gc
import platform
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Any, Dict, List
from unittest.mock import Mock, patch

import psutil
import pytest

from pyhub_office_automation.excel.utils import COMResourceManager
from pyhub_office_automation.excel.utils_timeout import execute_pivot_operation_with_cleanup, execute_with_timeout


class MemoryTracker:
    """메모리 사용량 추적 클래스"""

    def __init__(self):
        self.process = psutil.Process()
        self.initial_memory = self.get_memory_usage()
        self.peak_memory = self.initial_memory
        self.measurements = []

    def get_memory_usage(self) -> float:
        """현재 메모리 사용량(MB) 반환"""
        return self.process.memory_info().rss / 1024 / 1024

    def record_measurement(self, label: str = ""):
        """메모리 사용량 기록"""
        current_memory = self.get_memory_usage()
        self.peak_memory = max(self.peak_memory, current_memory)
        self.measurements.append((label, current_memory, time.time()))

    @property
    def memory_increase(self) -> float:
        """초기 대비 메모리 증가량(MB)"""
        return self.get_memory_usage() - self.initial_memory

    @property
    def peak_increase(self) -> float:
        """초기 대비 피크 메모리 증가량(MB)"""
        return self.peak_memory - self.initial_memory


class LargeCOMObject:
    """대용량 데이터를 가진 Mock COM 객체"""

    def __init__(self, size_mb: float = 1.0):
        self.name = f"LargeCOMObject_{size_mb}MB"
        self.data = bytearray(int(size_mb * 1024 * 1024))  # 지정된 크기의 데이터
        self.api = Mock()
        self.closed = False

    def close(self):
        self.closed = True
        # 대용량 데이터 해제
        self.data = None


class TestCOMMemoryManagement:
    """COM 메모리 관리 테스트"""

    @pytest.mark.skipif(platform.system() != "Windows", reason="COM은 Windows 전용")
    def test_memory_leak_prevention_basic(self):
        """기본 메모리 누수 방지 테스트"""
        tracker = MemoryTracker()
        tracker.record_measurement("시작")

        # 여러 번의 COM 객체 생성 및 정리
        for iteration in range(10):
            with COMResourceManager() as com_manager:
                # 중간 크기 객체들 생성
                for i in range(5):
                    large_obj = LargeCOMObject(0.5)  # 0.5MB 객체
                    com_manager.add(large_obj)

            tracker.record_measurement(f"반복_{iteration+1}_완료")

        tracker.record_measurement("완료")

        # 메모리 증가량이 합리적인 수준인지 확인 (5MB 미만)
        assert tracker.memory_increase < 5.0, f"메모리 증가량 과다: {tracker.memory_increase:.2f}MB"

    def test_memory_cleanup_verification(self):
        """메모리 정리 검증 테스트"""
        tracker = MemoryTracker()
        objects_created = []

        # 대량 객체 생성
        with COMResourceManager() as com_manager:
            for i in range(20):
                obj = LargeCOMObject(0.2)  # 0.2MB 객체
                objects_created.append(obj)
                com_manager.add(obj)

            tracker.record_measurement("객체_생성_완료")

        tracker.record_measurement("정리_완료")

        # 모든 객체가 정리되었는지 확인
        for obj in objects_created:
            assert obj.closed is True, f"객체 {obj.name}이 정리되지 않음"

        # 강제 가비지 컬렉션
        for _ in range(3):
            gc.collect()

        tracker.record_measurement("GC_완료")

    @patch("gc.collect")
    def test_garbage_collection_optimization(self, mock_gc):
        """가비지 컬렉션 최적화 테스트"""
        # 초기 호출 횟수
        initial_calls = mock_gc.call_count

        # 대량 COM 객체 처리
        with COMResourceManager() as com_manager:
            for i in range(50):
                obj = Mock()
                obj.close = Mock()
                com_manager.add(obj)

        # COMResourceManager가 적절한 횟수의 gc.collect를 호출했는지 확인
        total_calls = mock_gc.call_count - initial_calls
        assert total_calls == 3, f"예상 3회, 실제 {total_calls}회 gc.collect 호출"

    def test_concurrent_memory_management(self):
        """동시 실행 시 메모리 관리 테스트"""
        tracker = MemoryTracker()

        def create_com_objects(thread_id: int) -> str:
            """스레드별 COM 객체 생성 함수"""
            with COMResourceManager() as com_manager:
                objects = []
                for i in range(10):
                    obj = LargeCOMObject(0.1)  # 0.1MB 객체
                    objects.append(obj)
                    com_manager.add(obj)

                # 짧은 작업 시뮬레이션
                time.sleep(0.1)

                return f"Thread_{thread_id}_완료"

        # 동시에 5개 스레드 실행
        with ThreadPoolExecutor(max_workers=5) as executor:
            futures = [executor.submit(create_com_objects, i) for i in range(5)]
            results = [future.result() for future in as_completed(futures)]

        tracker.record_measurement("동시_작업_완료")

        # 모든 스레드가 성공적으로 완료
        assert len(results) == 5
        assert all("완료" in result for result in results)

        # 메모리 증가량이 합리적 수준인지 확인
        assert tracker.memory_increase < 10.0, f"동시 작업 후 메모리 증가량 과다: {tracker.memory_increase:.2f}MB"

    def test_memory_peak_monitoring(self):
        """메모리 피크 모니터링 테스트"""
        tracker = MemoryTracker()

        # 점진적으로 증가하는 메모리 사용량 테스트
        for batch_size in [5, 10, 20, 10, 5]:
            with COMResourceManager() as com_manager:
                for i in range(batch_size):
                    obj = LargeCOMObject(0.1)
                    com_manager.add(obj)

                tracker.record_measurement(f"배치_{batch_size}개")

        # 피크 메모리 증가량 확인
        print(f"피크 메모리 증가량: {tracker.peak_increase:.2f}MB")
        assert tracker.peak_increase < 15.0, f"피크 메모리 증가량 과다: {tracker.peak_increase:.2f}MB"


class TestTimeoutPerformance:
    """타임아웃 처리 성능 테스트"""

    def test_timeout_overhead_measurement(self):
        """타임아웃 처리 오버헤드 측정 테스트"""

        def quick_operation():
            return "quick_result"

        # 타임아웃 없이 직접 실행
        start_time = time.time()
        direct_result = quick_operation()
        direct_time = time.time() - start_time

        # 타임아웃과 함께 실행
        start_time = time.time()
        success, timeout_result, error = execute_with_timeout(quick_operation, timeout=10)
        timeout_time = time.time() - start_time

        # 결과 검증
        assert success is True
        assert timeout_result == direct_result

        # 오버헤드가 100ms 미만이어야 함
        overhead = timeout_time - direct_time
        assert overhead < 0.1, f"타임아웃 오버헤드 과다: {overhead*1000:.2f}ms"

    def test_multiple_timeout_operations_performance(self):
        """다중 타임아웃 작업 성능 테스트"""

        def simple_operation(value):
            time.sleep(0.01)  # 10ms 작업
            return f"result_{value}"

        start_time = time.time()

        # 100번의 타임아웃 작업 실행
        results = []
        for i in range(100):
            success, result, error = execute_with_timeout(simple_operation, args=(i,), timeout=5)
            results.append((success, result))

        total_time = time.time() - start_time

        # 모든 작업이 성공해야 함
        assert all(success for success, _ in results)

        # 전체 시간이 5초 미만이어야 함 (각 작업이 10ms이므로 충분히 가능)
        assert total_time < 5.0, f"100개 작업이 {total_time:.2f}초 소요 (너무 느림)"

    @patch("gc.collect")
    def test_cleanup_operation_performance(self, mock_gc):
        """정리 작업 성능 테스트"""

        def memory_intensive_operation():
            # 메모리 집약적 작업 시뮬레이션
            large_list = [i * "data" for i in range(1000)]
            return len(large_list)

        start_time = time.time()

        success, result, error = execute_pivot_operation_with_cleanup(
            memory_intensive_operation, timeout=10, description="메모리 집약적 작업"
        )

        execution_time = time.time() - start_time

        assert success is True
        assert result == 1000

        # 실행 시간이 1초 미만이어야 함
        assert execution_time < 1.0, f"정리 작업이 {execution_time:.2f}초 소요 (너무 느림)"

        # gc.collect가 적절히 호출되었는지 확인
        assert mock_gc.call_count >= 2


class TestScalabilityTests:
    """확장성 테스트"""

    def test_large_number_of_com_objects(self):
        """대량 COM 객체 처리 테스트"""
        tracker = MemoryTracker()
        object_count = 1000

        start_time = time.time()

        with COMResourceManager() as com_manager:
            objects = []
            for i in range(object_count):
                obj = Mock()
                obj.name = f"Object_{i}"
                obj.close = Mock()
                objects.append(obj)
                com_manager.add(obj)

            tracker.record_measurement("객체_생성_완료")

        processing_time = time.time() - start_time
        tracker.record_measurement("정리_완료")

        # 성능 검증
        assert processing_time < 5.0, f"{object_count}개 객체 처리가 {processing_time:.2f}초 소요"

        # 모든 객체가 정리되었는지 확인
        for obj in objects:
            obj.close.assert_called_once()

        # 메모리 사용량 확인
        assert tracker.memory_increase < 20.0, f"메모리 증가량 과다: {tracker.memory_increase:.2f}MB"

    def test_deep_nested_com_managers(self):
        """깊은 중첩 COM 매니저 테스트"""

        def create_nested_managers(depth: int) -> str:
            if depth <= 0:
                return "base_case"

            with COMResourceManager() as com_manager:
                obj = Mock()
                obj.close = Mock()
                obj.name = f"Object_depth_{depth}"
                com_manager.add(obj)

                result = create_nested_managers(depth - 1)
                return f"depth_{depth}_" + result

        start_time = time.time()
        result = create_nested_managers(20)  # 20단계 중첩
        processing_time = time.time() - start_time

        assert "depth_20_" in result
        assert "base_case" in result

        # 깊은 중첩도 빠르게 처리되어야 함
        assert processing_time < 1.0, f"20단계 중첩 처리가 {processing_time:.2f}초 소요"

    def test_high_frequency_operations(self):
        """고빈도 작업 테스트"""
        tracker = MemoryTracker()
        operation_count = 500

        def quick_com_operation(op_id: int):
            with COMResourceManager() as com_manager:
                obj = Mock()
                obj.close = Mock()
                com_manager.add(obj)
                return f"op_{op_id}_완료"

        start_time = time.time()

        results = []
        for i in range(operation_count):
            result = quick_com_operation(i)
            results.append(result)

            if i % 100 == 0:
                tracker.record_measurement(f"작업_{i}_완료")

        total_time = time.time() - start_time
        tracker.record_measurement("모든_작업_완료")

        # 성능 검증
        assert len(results) == operation_count
        assert total_time < 10.0, f"{operation_count}개 고빈도 작업이 {total_time:.2f}초 소요"

        # 평균 작업 시간
        avg_time_per_op = total_time / operation_count * 1000
        assert avg_time_per_op < 20.0, f"작업당 평균 {avg_time_per_op:.2f}ms (너무 느림)"

        # 메모리 누수 검증
        assert tracker.memory_increase < 10.0, f"고빈도 작업 후 메모리 증가량: {tracker.memory_increase:.2f}MB"


class TestMemoryLeakDetection:
    """메모리 누수 감지 테스트"""

    def test_repeated_com_operations_memory_stability(self):
        """반복적인 COM 작업의 메모리 안정성 테스트"""
        tracker = MemoryTracker()
        iterations = 50

        memory_snapshots = []

        for iteration in range(iterations):
            # COM 작업 수행
            with COMResourceManager() as com_manager:
                # 중간 크기 객체 생성
                for i in range(5):
                    obj = LargeCOMObject(0.2)
                    com_manager.add(obj)

                # 작업 시뮬레이션
                time.sleep(0.01)

            # 강제 가비지 컬렉션
            gc.collect()

            # 메모리 사용량 기록
            current_memory = tracker.get_memory_usage()
            memory_snapshots.append(current_memory)

            # 10회마다 메모리 체크
            if iteration % 10 == 9:
                tracker.record_measurement(f"반복_{iteration+1}")

        # 메모리 안정성 검증
        initial_memory = memory_snapshots[5]  # 처음 몇 개는 워밍업으로 제외
        final_memory = memory_snapshots[-1]
        memory_growth = final_memory - initial_memory

        assert memory_growth < 5.0, f"메모리 지속적 증가 감지: {memory_growth:.2f}MB"

        # 메모리 사용량의 표준편차가 작아야 함 (안정성 지표)
        import statistics

        memory_std = statistics.stdev(memory_snapshots[10:])  # 워밍업 제외
        assert memory_std < 2.0, f"메모리 사용량 불안정: 표준편차 {memory_std:.2f}MB"

    def test_circular_reference_cleanup(self):
        """순환 참조 정리 테스트"""
        tracker = MemoryTracker()

        def create_circular_references():
            with COMResourceManager() as com_manager:
                objects = []

                # 순환 참조 생성
                for i in range(10):
                    obj = Mock()
                    obj.name = f"CircularObject_{i}"
                    obj.close = Mock()
                    objects.append(obj)
                    com_manager.add(obj)

                # 순환 참조 설정
                for i, obj in enumerate(objects):
                    obj.next_obj = objects[(i + 1) % len(objects)]
                    obj.prev_obj = objects[i - 1]

                tracker.record_measurement("순환_참조_생성")

        # 반복 실행
        for iteration in range(10):
            create_circular_references()
            tracker.record_measurement(f"순환_참조_정리_{iteration}")

        # 메모리 증가량이 합리적인지 확인
        assert tracker.memory_increase < 5.0, f"순환 참조 정리 후 메모리 증가: {tracker.memory_increase:.2f}MB"

    def test_timeout_operation_memory_leak(self):
        """타임아웃 작업의 메모리 누수 테스트"""
        tracker = MemoryTracker()

        def memory_consuming_operation():
            # 메모리를 사용하는 작업
            data = [f"data_{i}" * 1000 for i in range(100)]
            time.sleep(0.1)
            return len(data)

        # 반복적인 타임아웃 작업
        for iteration in range(20):
            success, result, error = execute_pivot_operation_with_cleanup(
                memory_consuming_operation, timeout=5, description=f"메모리_작업_{iteration}"
            )

            assert success is True

            if iteration % 5 == 4:
                tracker.record_measurement(f"타임아웃_작업_{iteration+1}")

        # 메모리 누수 검증
        assert tracker.memory_increase < 10.0, f"타임아웃 작업 후 메모리 증가: {tracker.memory_increase:.2f}MB"

    @patch("gc.collect")
    def test_gc_collect_effectiveness(self, mock_gc):
        """gc.collect 호출 효과 테스트"""
        # 초기 메모리 상태
        initial_objects = len(gc.get_objects())

        def create_temporary_objects():
            temp_objects = []
            for i in range(100):
                obj = Mock()
                obj.large_data = "x" * 10000
                temp_objects.append(obj)

            return temp_objects

        # 임시 객체 생성
        objects = create_temporary_objects()

        with COMResourceManager() as com_manager:
            for obj in objects:
                com_manager.add(obj)

        # gc.collect가 호출되었는지 확인
        assert mock_gc.call_count >= 3

        # 실제 gc.collect 실행하여 정리 확인
        gc.collect()
        final_objects = len(gc.get_objects())

        # 객체 수가 크게 증가하지 않았는지 확인
        object_increase = final_objects - initial_objects
        assert object_increase < 50, f"정리 후 객체 수 증가: {object_increase}"


if __name__ == "__main__":
    # 성능 테스트 실행 예시
    print("COM 리소스 관리 성능 테스트 실행 중...")

    # 기본 메모리 누수 테스트
    test_memory = TestCOMMemoryManagement()
    test_memory.test_memory_leak_prevention_basic()

    print("성능 테스트 완료!")
