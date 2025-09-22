"""
타임아웃 처리를 위한 유틸리티 함수들
COM API 타임아웃 문제 해결을 위한 헬퍼 함수 제공
"""

import threading
from typing import Any, Callable, Optional, Tuple


def execute_with_timeout(
    func: Callable, args: tuple = (), kwargs: dict = None, timeout: int = 10
) -> Tuple[bool, Any, Optional[str]]:
    """
    함수를 타임아웃과 함께 실행합니다.

    Args:
        func: 실행할 함수
        args: 함수 인자 (튜플)
        kwargs: 함수 키워드 인자 (딕셔너리)
        timeout: 타임아웃 시간 (초, 기본값: 10)

    Returns:
        (success: bool, result: Any, error_msg: Optional[str])
        - success: 성공 여부
        - result: 함수 실행 결과 (성공 시) 또는 None (실패 시)
        - error_msg: 에러 메시지 (실패 시) 또는 None (성공 시)
    """
    if kwargs is None:
        kwargs = {}

    result = [None]
    exception = [None]

    def target():
        try:
            result[0] = func(*args, **kwargs)
        except Exception as e:
            exception[0] = e

    thread = threading.Thread(target=target)
    thread.daemon = True
    thread.start()
    thread.join(timeout)

    if thread.is_alive():
        # 타임아웃 발생
        return False, None, f"작업이 {timeout}초 내에 완료되지 않아 타임아웃되었습니다"

    if exception[0] is not None:
        # 함수 실행 중 예외 발생
        return False, None, str(exception[0])

    # 성공
    return True, result[0], None


def try_pivot_layout_connection(chart, pivot_table, timeout: int = 10) -> Tuple[bool, Optional[str]]:
    """
    피벗차트 연결을 시도하고 타임아웃 시 실패를 반환합니다.

    Args:
        chart: Excel 차트 객체
        pivot_table: Excel 피벗테이블 객체
        timeout: 타임아웃 시간 (초, 기본값: 10)

    Returns:
        (success: bool, error_msg: Optional[str])
    """

    def set_pivot_layout():
        chart.PivotLayout.PivotTable = pivot_table
        return True

    success, _, error_msg = execute_with_timeout(set_pivot_layout, timeout=timeout)

    if not success:
        if "타임아웃" in str(error_msg):
            return False, "피벗차트 연결이 타임아웃되었습니다. 정적 차트로 생성됩니다."
        else:
            return False, f"피벗차트 연결 실패: {error_msg}"

    return True, None
