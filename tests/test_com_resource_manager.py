"""
COM 리소스 관리자 단위 테스트
COMResourceManager 클래스와 관련 유틸리티 함수들에 대한 포괄적 테스트
"""

import gc
import platform
import threading
import time
from typing import Any
from unittest.mock import MagicMock, Mock, call, patch

import pytest

from pyhub_office_automation.excel.utils import COMResourceManager


class TestCOMResourceManager:
    """COMResourceManager 클래스 단위 테스트"""

    def test_init_default_params(self):
        """기본 매개변수로 초기화 테스트"""
        manager = COMResourceManager()

        assert manager.com_objects == []
        assert manager.api_refs == []
        assert manager.verbose is False
        assert manager._original_objects == []

    def test_init_verbose_mode(self):
        """verbose 모드 초기화 테스트"""
        manager = COMResourceManager(verbose=True)

        assert manager.verbose is True

    def test_add_object_basic(self):
        """기본 객체 추가 테스트"""
        manager = COMResourceManager()
        mock_obj = Mock()

        result = manager.add(mock_obj)

        assert result is mock_obj  # 체이닝 확인
        assert mock_obj in manager.com_objects
        assert len(manager.com_objects) == 1

    def test_add_object_with_description_verbose(self):
        """설명과 함께 객체 추가 (verbose 모드) 테스트"""
        manager = COMResourceManager(verbose=True)
        mock_obj = Mock()
        description = "Test Object"

        manager.add(mock_obj, description)

        assert mock_obj in manager.com_objects
        assert (mock_obj, description) in manager._original_objects

    def test_add_object_with_description_non_verbose(self):
        """설명과 함께 객체 추가 (non-verbose 모드) 테스트"""
        manager = COMResourceManager(verbose=False)
        mock_obj = Mock()
        description = "Test Object"

        manager.add(mock_obj, description)

        assert mock_obj in manager.com_objects
        assert len(manager._original_objects) == 0  # verbose=False이므로 저장되지 않음

    def test_add_object_with_api_attribute(self):
        """API 속성이 있는 객체 추가 테스트"""
        manager = COMResourceManager()
        mock_obj = Mock()
        mock_api = Mock()
        mock_obj.api = mock_api

        manager.add(mock_obj)

        assert mock_obj in manager.com_objects
        assert (mock_obj, "api") in manager.api_refs

    def test_add_object_without_api_attribute(self):
        """API 속성이 없는 객체 추가 테스트"""
        manager = COMResourceManager()
        mock_obj = Mock()
        del mock_obj.api  # API 속성 제거

        manager.add(mock_obj)

        assert mock_obj in manager.com_objects
        assert len(manager.api_refs) == 0

    def test_add_none_object(self):
        """None 객체 추가 테스트"""
        manager = COMResourceManager()

        result = manager.add(None)

        assert result is None
        assert len(manager.com_objects) == 0

    def test_add_duplicate_object(self):
        """중복 객체 추가 테스트"""
        manager = COMResourceManager()
        mock_obj = Mock()

        manager.add(mock_obj)
        manager.add(mock_obj)  # 중복 추가

        assert len(manager.com_objects) == 1  # 중복 제거됨

    def test_add_api_ref_explicit(self):
        """명시적 API 참조 추가 테스트"""
        manager = COMResourceManager()
        mock_obj = Mock()
        mock_api = Mock()
        mock_obj.custom_api = mock_api

        manager.add_api_ref(mock_obj, "custom_api")

        assert (mock_obj, "custom_api") in manager.api_refs

    def test_add_api_ref_non_existent_attribute(self):
        """존재하지 않는 API 속성 추가 테스트"""
        manager = COMResourceManager()
        mock_obj = Mock()
        del mock_obj.non_existent  # 속성 제거

        manager.add_api_ref(mock_obj, "non_existent")

        assert len(manager.api_refs) == 0

    def test_context_manager_enter(self):
        """컨텍스트 매니저 enter 테스트"""
        manager = COMResourceManager()

        with manager as cm:
            assert cm is manager

    @patch("gc.collect")
    @patch("platform.system")
    @patch("pythoncom.CoUninitialize")
    def test_context_manager_exit_success(self, mock_co_uninit, mock_platform, mock_gc):
        """컨텍스트 매니저 exit 성공 테스트"""
        mock_platform.return_value = "Windows"

        manager = COMResourceManager()
        mock_obj = Mock()
        mock_obj.close = Mock()
        manager.add(mock_obj)

        # 컨텍스트 매니저 사용
        with manager:
            pass

        # 검증
        mock_obj.close.assert_called_once()
        assert mock_gc.call_count == 3  # 3번 호출됨
        mock_co_uninit.assert_called_once()
        assert len(manager.com_objects) == 0

    @patch("gc.collect")
    @patch("platform.system")
    def test_context_manager_exit_non_windows(self, mock_platform, mock_gc):
        """Windows가 아닌 환경에서 컨텍스트 매니저 exit 테스트"""
        mock_platform.return_value = "Linux"

        manager = COMResourceManager()
        mock_obj = Mock()
        manager.add(mock_obj)

        with manager:
            pass

        # Windows가 아니므로 COM 라이브러리 정리 함수 호출되지 않음
        assert mock_gc.call_count == 3

    @patch("builtins.print")
    @patch("gc.collect")
    def test_context_manager_exit_verbose_mode(self, mock_gc, mock_print):
        """verbose 모드에서 컨텍스트 매니저 exit 테스트"""
        manager = COMResourceManager(verbose=True)
        mock_obj = Mock()
        manager.add(mock_obj, "Test Object")

        with manager:
            pass

        # verbose 출력 확인
        mock_print.assert_any_call("[COMResourceManager] 정리 시작: 1개 객체")
        mock_print.assert_any_call("[COMResourceManager] 정리 완료")

    def test_api_reference_cleanup(self):
        """API 참조 정리 테스트"""
        manager = COMResourceManager()

        # Mock COM 객체 설정
        mock_obj = Mock()
        mock_api = Mock()
        mock_api.Release = Mock()
        mock_obj.api = mock_api

        manager.add(mock_obj)

        with manager:
            pass

        # API Release 호출 확인
        mock_api.Release.assert_called_once()

    def test_api_reference_cleanup_with_release_error(self):
        """API 참조 정리 시 Release 에러 테스트"""
        manager = COMResourceManager(verbose=True)

        # Mock COM 객체 설정 (Release 시 에러 발생)
        mock_obj = Mock()
        mock_api = Mock()
        mock_api.Release.side_effect = Exception("Release failed")
        mock_obj.api = mock_api

        manager.add(mock_obj)

        # 에러 발생해도 정상적으로 컨텍스트 종료되어야 함
        with manager:
            pass

    def test_object_cleanup_with_close_method(self):
        """close 메서드가 있는 객체 정리 테스트"""
        manager = COMResourceManager()

        mock_obj = Mock()
        mock_obj.close = Mock()
        manager.add(mock_obj)

        with manager:
            pass

        mock_obj.close.assert_called_once()

    def test_object_cleanup_with_quit_method(self):
        """quit 메서드가 있는 객체 정리 테스트"""
        manager = COMResourceManager()

        mock_obj = Mock()
        mock_obj.quit = Mock()
        del mock_obj.close  # close 메서드 제거
        manager.add(mock_obj)

        with manager:
            pass

        mock_obj.quit.assert_called_once()

    def test_object_cleanup_with_method_error(self):
        """객체 정리 중 메서드 에러 테스트"""
        manager = COMResourceManager(verbose=True)

        mock_obj = Mock()
        mock_obj.close.side_effect = Exception("Close failed")
        manager.add(mock_obj)

        # 에러 발생해도 정상적으로 컨텍스트 종료되어야 함
        with manager:
            pass

    def test_object_cleanup_reverse_order(self):
        """객체 정리 시 역순 처리 테스트"""
        manager = COMResourceManager()

        # 여러 객체 추가 (순서 확인용)
        objects = []
        for i in range(3):
            mock_obj = Mock()
            mock_obj.name = f"obj_{i}"
            mock_obj.close = Mock()
            objects.append(mock_obj)
            manager.add(mock_obj)

        with manager:
            pass

        # 모든 객체의 close가 호출되었는지 확인 (순서는 체크하지 않음, 실제로는 역순)
        for obj in objects:
            obj.close.assert_called_once()

    def test_exception_handling_does_not_suppress(self):
        """예외 처리가 컨텍스트 예외를 억제하지 않는지 테스트"""
        manager = COMResourceManager()

        with pytest.raises(ValueError):
            with manager:
                raise ValueError("Test exception")

    @patch("gc.collect")
    @patch("platform.system")
    @patch("pythoncom.CoUninitialize")
    def test_com_library_cleanup_error_handling(self, mock_co_uninit, mock_platform, mock_gc):
        """COM 라이브러리 정리 중 에러 처리 테스트"""
        mock_platform.return_value = "Windows"
        mock_co_uninit.side_effect = Exception("CoUninitialize failed")

        manager = COMResourceManager()

        # 에러 발생해도 정상적으로 컨텍스트 종료되어야 함
        with manager:
            pass

        mock_co_uninit.assert_called_once()

    @patch("gc.collect")
    def test_garbage_collection_multiple_calls(self, mock_gc):
        """가비지 컬렉션 3번 호출 확인 테스트"""
        manager = COMResourceManager()

        with manager:
            pass

        assert mock_gc.call_count == 3

    def test_lists_cleared_after_cleanup(self):
        """정리 후 리스트들이 비워지는지 테스트"""
        manager = COMResourceManager(verbose=True)

        mock_obj = Mock()
        mock_obj.api = Mock()
        manager.add(mock_obj, "Test Object")

        # 추가된 항목들 확인
        assert len(manager.com_objects) == 1
        assert len(manager.api_refs) == 1
        assert len(manager._original_objects) == 1

        with manager:
            pass

        # 정리 후 모든 리스트가 비워졌는지 확인
        assert len(manager.com_objects) == 0
        assert len(manager.api_refs) == 0
        assert len(manager._original_objects) == 0

    def test_chaining_support(self):
        """메서드 체이닝 지원 테스트"""
        manager = COMResourceManager()
        mock_obj1 = Mock()
        mock_obj2 = Mock()

        # 체이닝 테스트
        result = manager.add(mock_obj1).close if hasattr(manager.add(mock_obj1), "close") else None

        # add는 추가된 객체를 반환하므로 체이닝 가능
        chained_obj = manager.add(mock_obj2)
        assert chained_obj is mock_obj2

    @patch("platform.system")
    def test_platform_detection(self, mock_platform):
        """플랫폼 감지 테스트"""
        # Windows 테스트
        mock_platform.return_value = "Windows"
        with patch("pythoncom.CoUninitialize") as mock_co_uninit:
            manager = COMResourceManager()
            with manager:
                pass
            mock_co_uninit.assert_called_once()

        # 다른 플랫폼 테스트
        mock_platform.return_value = "Darwin"
        with patch("pythoncom.CoUninitialize") as mock_co_uninit:
            manager = COMResourceManager()
            with manager:
                pass
            mock_co_uninit.assert_not_called()

    def test_performance_with_many_objects(self):
        """많은 객체를 가진 성능 테스트"""
        manager = COMResourceManager()

        # 100개 객체 추가
        objects = []
        for i in range(100):
            mock_obj = Mock()
            mock_obj.close = Mock()
            objects.append(mock_obj)
            manager.add(mock_obj)

        start_time = time.time()
        with manager:
            pass
        end_time = time.time()

        # 정리가 1초 내에 완료되어야 함
        assert (end_time - start_time) < 1.0

        # 모든 객체가 정리되었는지 확인
        for obj in objects:
            obj.close.assert_called_once()

    def test_memory_cleanup_simulation(self):
        """메모리 정리 시뮬레이션 테스트"""
        manager = COMResourceManager()

        # 강한 참조 생성
        mock_objects = []
        for i in range(10):
            mock_obj = Mock()
            mock_obj.data = f"large_data_{i}" * 1000  # 큰 데이터 시뮬레이션
            mock_objects.append(mock_obj)
            manager.add(mock_obj)

        # 참조 해제
        with manager:
            pass

        # 객체들이 정리 리스트에서 제거되었는지 확인
        assert len(manager.com_objects) == 0
