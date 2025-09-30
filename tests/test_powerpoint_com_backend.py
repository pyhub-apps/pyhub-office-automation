"""
PowerPoint COM 백엔드 기본 테스트
Windows 환경에서만 실행됨
"""

import platform

import pytest

# Windows 전용 테스트
pytestmark = pytest.mark.skipif(platform.system() != "Windows", reason="COM 백엔드는 Windows 전용")


class TestBackendSelector:
    """백엔드 선택 로직 테스트"""

    def test_detect_backend_auto(self):
        """자동 백엔드 감지 테스트"""
        from pyhub_office_automation.powerpoint.backend_selector import detect_backend

        backend = detect_backend()
        assert backend in ["com", "python-pptx"]

        # Windows에서는 COM이 우선
        if platform.system() == "Windows":
            try:
                import win32com.client

                assert backend == "com"
            except ImportError:
                assert backend == "python-pptx"

    def test_detect_backend_force_com(self):
        """COM 강제 선택 테스트"""
        from pyhub_office_automation.powerpoint.backend_selector import detect_backend

        if platform.system() != "Windows":
            with pytest.raises(RuntimeError):
                detect_backend("com")
        else:
            try:
                import win32com.client

                backend = detect_backend("com")
                assert backend == "com"
            except ImportError:
                with pytest.raises(RuntimeError):
                    detect_backend("com")

    def test_check_backend_availability(self):
        """백엔드 가용성 체크 테스트"""
        from pyhub_office_automation.powerpoint.backend_selector import check_backend_availability

        # python-pptx는 항상 체크 가능
        available, error = check_backend_availability("python-pptx")
        assert isinstance(available, bool)
        if not available:
            assert error is not None

        # COM은 Windows + pywin32 필요
        available, error = check_backend_availability("com")
        assert isinstance(available, bool)
        if not available:
            assert error is not None

    def test_get_backend_info(self):
        """백엔드 정보 조회 테스트"""
        from pyhub_office_automation.powerpoint.backend_selector import get_backend_info

        # COM 정보
        info = get_backend_info("com")
        assert info["backend"] == "com"
        assert "Windows" in info["platform"]
        assert len(info["features"]) > 0

        # python-pptx 정보
        info = get_backend_info("python-pptx")
        assert info["backend"] == "python-pptx"
        assert len(info["limitations"]) > 0


class TestPowerPointCOM:
    """PowerPointCOM 클래스 기본 테스트"""

    def test_import_com_backend(self):
        """COM 백엔드 임포트 테스트"""
        try:
            from pyhub_office_automation.powerpoint.com_backend import PowerPointCOM

            assert PowerPointCOM is not None
        except RuntimeError as e:
            # Windows가 아니면 예상된 에러
            if platform.system() != "Windows":
                assert "Windows에서만" in str(e)
            else:
                raise

    @pytest.mark.skipif(platform.system() != "Windows", reason="COM 백엔드는 Windows 전용")
    def test_powerpoint_com_init(self):
        """PowerPointCOM 초기화 테스트"""
        try:
            import win32com.client
        except ImportError:
            pytest.skip("pywin32가 설치되지 않음")

        from pyhub_office_automation.powerpoint.com_backend import PowerPointCOM

        # PowerPoint COM 초기화 (항상 visible=True)
        ppt = PowerPointCOM()
        assert ppt.app is not None

        # 정리
        ppt.close(quit_app=False)


class TestUtilsFunctions:
    """utils.py 함수 테스트"""

    def test_get_powerpoint_backend(self):
        """get_powerpoint_backend 함수 테스트"""
        from pyhub_office_automation.powerpoint.utils import get_powerpoint_backend

        backend = get_powerpoint_backend()
        assert backend in ["com", "python-pptx"]

    def test_get_powerpoint_backend_force(self):
        """강제 백엔드 지정 테스트"""
        from pyhub_office_automation.powerpoint.utils import get_powerpoint_backend

        # python-pptx는 항상 가능해야 함 (설치되어 있다면)
        try:
            import pptx

            backend = get_powerpoint_backend(force_backend="python-pptx")
            assert backend == "python-pptx"
        except ImportError:
            # python-pptx가 없으면 에러 발생 예상
            with pytest.raises(RuntimeError):
                get_powerpoint_backend(force_backend="python-pptx")

    def test_create_success_response(self):
        """성공 응답 생성 테스트"""
        from pyhub_office_automation.powerpoint.utils import create_success_response

        result = create_success_response(command="test-command", data={"test": "data"}, message="Test message")

        assert result["success"] is True
        assert result["command"] == "test-command"
        assert result["data"] == {"test": "data"}
        assert result["message"] == "Test message"
        assert "version" in result

    def test_create_error_response(self):
        """에러 응답 생성 테스트"""
        from pyhub_office_automation.powerpoint.utils import create_error_response

        result = create_error_response(command="test-command", error="Test error", error_type="TestError")

        assert result["success"] is False
        assert result["command"] == "test-command"
        assert result["error"] == "Test error"
        assert result["error_type"] == "TestError"
        assert "version" in result


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
