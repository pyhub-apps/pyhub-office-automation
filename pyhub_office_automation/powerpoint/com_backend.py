"""
PowerPoint COM 백엔드 (Windows 전용)
pywin32를 사용한 PowerPoint 완전 제어
"""

import gc
import platform
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

from pyhub_office_automation.version import get_version


class PowerPointCOM:
    """
    PowerPoint COM 자동화 래퍼 클래스

    Windows pywin32를 사용하여 PowerPoint 애플리케이션을 완전히 제어합니다.
    Excel xlwings 패턴을 참고하되, PowerPoint 특화 기능 구현.

    Features:
    - PowerPoint 앱 생명주기 관리
    - Presentation 열기/생성/저장
    - COM 리소스 자동 정리
    - 에러 핸들링 및 타임아웃 처리

    Example:
        >>> ppt = PowerPointCOM(visible=True)
        >>> prs = ppt.open_presentation("report.pptx")
        >>> ppt.close()
    """

    def __init__(self, add_if_not_running: bool = True):
        """
        PowerPoint COM 객체를 초기화합니다.

        Note:
            PowerPoint는 Excel과 달리 항상 visible=True로 실행됩니다.
            이는 PowerPoint COM API의 제약사항입니다.

        Args:
            add_if_not_running: PowerPoint가 실행 중이 아니면 새로 시작 (기본: True)

        Raises:
            ImportError: pywin32가 설치되지 않은 경우
            RuntimeError: Windows가 아닌 플랫폼에서 실행한 경우
        """
        # 플랫폼 체크
        if platform.system() != "Windows":
            raise RuntimeError("PowerPoint COM 백엔드는 Windows에서만 사용 가능합니다")

        # pywin32 체크
        try:
            import pythoncom
            import win32com.client

            self._win32com = win32com
            self._pythoncom = pythoncom
        except ImportError:
            raise ImportError("pywin32 패키지가 설치되지 않았습니다. " "'pip install pywin32'로 설치하세요")

        self._app = None
        self._presentations = {}  # {name: presentation} 매핑

        # PowerPoint 애플리케이션 초기화
        self._initialize_app(add_if_not_running)

    def _initialize_app(self, add_if_not_running: bool = True):
        """PowerPoint 애플리케이션을 초기화합니다."""
        try:
            # 기존 실행 중인 PowerPoint에 연결 시도
            self._app = self._win32com.client.GetActiveObject("PowerPoint.Application")
        except Exception:
            if add_if_not_running:
                # 실행 중이 아니면 새로 시작
                self._app = self._win32com.client.Dispatch("PowerPoint.Application")
            else:
                raise RuntimeError("PowerPoint가 실행 중이지 않습니다")

        # PowerPoint는 항상 visible=True여야 함 (COM 제약사항)
        # visible=False 설정은 무시됨
        try:
            self._app.Visible = 1
        except Exception:
            # Visible 설정 실패해도 계속 진행
            pass

    @property
    def app(self):
        """PowerPoint Application COM 객체를 반환합니다."""
        return self._app

    @property
    def presentations(self):
        """현재 관리 중인 Presentation 객체들을 반환합니다."""
        return self._presentations

    def get_or_create_app(self):
        """
        PowerPoint 애플리케이션을 가져오거나 생성합니다.

        Returns:
            PowerPoint Application COM 객체
        """
        if self._app is None:
            self._initialize_app()
        return self._app

    def open_presentation(self, file_path: Union[str, Path], read_only: bool = False, with_window: bool = True) -> Any:
        """
        PowerPoint 프레젠테이션을 엽니다.

        Args:
            file_path: 프레젠테이션 파일 경로
            read_only: 읽기 전용으로 열기 (기본: False)
            with_window: 윈도우와 함께 열기 (기본: True)

        Returns:
            Presentation COM 객체

        Raises:
            FileNotFoundError: 파일이 존재하지 않는 경우
        """
        # 경로 정규화
        file_path = Path(file_path).resolve()

        if not file_path.exists():
            raise FileNotFoundError(f"프레젠테이션 파일을 찾을 수 없습니다: {file_path}")

        # 프레젠테이션 열기
        # Parameters: FileName, ReadOnly, Untitled, WithWindow
        prs = self._app.Presentations.Open(
            str(file_path), int(read_only), 0, int(with_window)  # ReadOnly  # Untitled (항상 0)  # WithWindow
        )

        # 관리 목록에 추가
        self._presentations[prs.Name] = prs

        return prs

    def create_presentation(self, save_path: Optional[Union[str, Path]] = None) -> Any:
        """
        새 PowerPoint 프레젠테이션을 생성합니다.

        Args:
            save_path: 저장 경로 (선택, 지정 시 즉시 저장)

        Returns:
            Presentation COM 객체
        """
        # 새 프레젠테이션 생성
        prs = self._app.Presentations.Add(WithWindow=1)

        # 저장 경로가 지정되면 즉시 저장
        if save_path:
            save_path = Path(save_path).resolve()
            prs.SaveAs(str(save_path))

        # 관리 목록에 추가
        self._presentations[prs.Name] = prs

        return prs

    def get_presentation_by_name(self, name: str) -> Optional[Any]:
        """
        이름으로 열려있는 프레젠테이션을 찾습니다.

        Args:
            name: 프레젠테이션 이름 (예: "report.pptx")

        Returns:
            Presentation COM 객체 또는 None
        """
        # 관리 목록에서 찾기
        if name in self._presentations:
            return self._presentations[name]

        # PowerPoint 앱에서 직접 찾기
        try:
            for prs in self._app.Presentations:
                if prs.Name == name:
                    self._presentations[name] = prs
                    return prs
        except Exception:
            pass

        return None

    def get_active_presentation(self) -> Optional[Any]:
        """
        현재 활성화된 프레젠테이션을 가져옵니다.

        Returns:
            Presentation COM 객체 또는 None
        """
        try:
            if self._app.Presentations.Count > 0:
                # ActivePresentation 속성 사용
                active_prs = self._app.ActivePresentation

                # 관리 목록에 추가
                if active_prs.Name not in self._presentations:
                    self._presentations[active_prs.Name] = active_prs

                return active_prs
        except Exception:
            # 활성 프레젠테이션이 없는 경우
            pass

        return None

    def list_presentations(self) -> List[Dict[str, Any]]:
        """
        현재 열려있는 모든 프레젠테이션 목록을 반환합니다.

        Returns:
            프레젠테이션 정보 리스트
            [{"name": "report.pptx", "path": "C:/...", "slide_count": 10, "saved": True}]
        """
        presentations = []

        try:
            for prs in self._app.Presentations:
                try:
                    info = {
                        "name": prs.Name,
                        "path": prs.FullName if prs.Path else None,
                        "slide_count": prs.Slides.Count,
                        "saved": bool(prs.Saved),
                    }
                    presentations.append(info)

                    # 관리 목록 업데이트
                    if prs.Name not in self._presentations:
                        self._presentations[prs.Name] = prs

                except Exception as e:
                    # 개별 프레젠테이션 정보 수집 실패
                    presentations.append({"name": getattr(prs, "Name", "Unknown"), "error": str(e)})
        except Exception:
            # Presentations 접근 실패
            pass

        return presentations

    def save_presentation(self, presentation: Any, file_path: Optional[Union[str, Path]] = None):
        """
        프레젠테이션을 저장합니다.

        Args:
            presentation: Presentation COM 객체
            file_path: 저장 경로 (선택, 미지정 시 원본 경로에 저장)
        """
        if file_path:
            # 다른 이름으로 저장
            file_path = Path(file_path).resolve()
            presentation.SaveAs(str(file_path))
        else:
            # 원본에 저장
            presentation.Save()

    def close_presentation(self, presentation: Any, save_changes: bool = True):
        """
        프레젠테이션을 닫습니다.

        Args:
            presentation: Presentation COM 객체
            save_changes: 변경사항 저장 여부 (기본: True)
        """
        try:
            # 관리 목록에서 제거
            if presentation.Name in self._presentations:
                del self._presentations[presentation.Name]

            # 프레젠테이션 닫기
            presentation.Close()

            # 저장 여부 처리는 PowerPoint가 자동으로 처리
            # (Saved 속성이 False면 사용자에게 물어봄)

        except Exception:
            # 이미 닫힌 경우 등 무시
            pass

    def close(self, quit_app: bool = False):
        """
        PowerPoint COM 객체를 정리합니다.

        Args:
            quit_app: PowerPoint 애플리케이션 종료 여부 (기본: False)
        """
        try:
            # 관리 중인 프레젠테이션 모두 정리
            self._presentations.clear()

            # 앱 종료 요청 시
            if quit_app and self._app:
                try:
                    self._app.Quit()
                except Exception:
                    pass

            # COM 객체 해제
            self._app = None

            # 가비지 컬렉션
            gc.collect()

        except Exception:
            # 정리 중 에러 무시
            pass

    def __enter__(self):
        """컨텍스트 매니저 진입"""
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """컨텍스트 매니저 종료"""
        self.close(quit_app=False)

    def __del__(self):
        """소멸자 - 리소스 정리"""
        try:
            self.close(quit_app=False)
        except Exception:
            pass


def get_or_open_presentation_com(
    file_path: Optional[str] = None, presentation_name: Optional[str] = None, create_if_not_found: bool = False
) -> tuple:
    """
    파일 경로 또는 프레젠테이션 이름으로 프레젠테이션을 가져오거나 엽니다.

    Args:
        file_path: 프레젠테이션 파일 경로
        presentation_name: 열려있는 프레젠테이션 이름
        create_if_not_found: 찾을 수 없을 때 새로 생성 (기본: False)

    Returns:
        (PowerPointCOM 인스턴스, Presentation COM 객체) 튜플

    Raises:
        ValueError: file_path와 presentation_name이 모두 없거나 둘 다 있는 경우
        FileNotFoundError: 파일이 존재하지 않는 경우

    Example:
        >>> ppt, prs = get_or_open_presentation_com(file_path="report.pptx")
        >>> # ... 작업 수행 ...
        >>> ppt.close()
    """
    # 입력 검증
    if file_path and presentation_name:
        raise ValueError("file_path와 presentation_name 중 하나만 지정해야 합니다")

    if not file_path and not presentation_name and not create_if_not_found:
        raise ValueError("file_path 또는 presentation_name 중 하나는 필수입니다")

    # PowerPoint COM 초기화
    ppt = PowerPointCOM()

    try:
        # 파일 경로로 열기
        if file_path:
            prs = ppt.open_presentation(file_path)
            return ppt, prs

        # 이름으로 찾기
        elif presentation_name:
            prs = ppt.get_presentation_by_name(presentation_name)
            if prs:
                return ppt, prs
            elif create_if_not_found:
                prs = ppt.create_presentation()
                return ppt, prs
            else:
                raise FileNotFoundError(f"열려있는 프레젠테이션을 찾을 수 없습니다: {presentation_name}")

        # 아무것도 없으면 활성 프레젠테이션 또는 새로 생성
        else:
            prs = ppt.get_active_presentation()
            if prs:
                return ppt, prs
            elif create_if_not_found:
                prs = ppt.create_presentation()
                return ppt, prs
            else:
                raise ValueError("활성 프레젠테이션이 없습니다")

    except Exception as e:
        # 에러 발생 시 COM 정리
        ppt.close()
        raise e
