"""
Python 환경 자동 탐지 유틸리티
- 시스템에 설치된 Python 인터프리터 자동 감지
- 경로와 버전 정보 추출
- 크로스 플랫폼 지원 (Windows, macOS, Linux)
"""

import os
import sys
import subprocess
import shutil
from pathlib import Path
from typing import List, Optional, Dict, Tuple
import re


class PythonInfo:
    """Python 인터프리터 정보를 담는 클래스"""

    def __init__(self, path: str, version: str, is_recommended: bool = False):
        self.path = path
        self.version = version
        self.is_recommended = is_recommended

    def __str__(self):
        return f"Python {self.version} at {self.path}"

    def to_dict(self) -> Dict:
        return {
            "path": self.path,
            "version": self.version,
            "is_recommended": self.is_recommended
        }


class PythonDetector:
    """Python 설치 자동 탐지 클래스"""

    # Python 3.13+ 권장
    RECOMMENDED_VERSION = (3, 13, 0)

    def __init__(self):
        self.found_pythons: List[PythonInfo] = []

    def detect_all(self) -> List[PythonInfo]:
        """모든 Python 설치를 탐지합니다."""
        self.found_pythons = []

        # 1. PATH 환경변수에서 찾기
        self._detect_from_path()

        # 2. 일반적인 설치 경로에서 찾기
        self._detect_from_common_paths()

        # 3. 중복 제거 및 정렬
        self._remove_duplicates()
        self._sort_by_version()

        # 4. 권장 버전 표시
        self._mark_recommended()

        return self.found_pythons

    def get_best_python(self) -> Optional[PythonInfo]:
        """가장 적합한 Python을 반환합니다."""
        if not self.found_pythons:
            self.detect_all()

        # 권장 버전이 있으면 우선 반환
        for python in self.found_pythons:
            if python.is_recommended:
                return python

        # 권장 버전이 없으면 가장 높은 버전 반환
        if self.found_pythons:
            return self.found_pythons[0]

        return None

    def _detect_from_path(self):
        """PATH 환경변수에서 Python 찾기"""
        python_names = ["python", "python3"]

        if sys.platform == "win32":
            python_names.extend(["python.exe", "python3.exe"])

        for name in python_names:
            python_path = shutil.which(name)
            if python_path:
                info = self._get_python_info(python_path)
                if info:
                    self.found_pythons.append(info)

    def _detect_from_common_paths(self):
        """일반적인 설치 경로에서 Python 찾기"""
        if sys.platform == "win32":
            self._detect_windows_paths()
        else:
            self._detect_unix_paths()

    def _detect_windows_paths(self):
        """Windows 일반 경로에서 Python 찾기"""
        # Python.org 설치 경로
        python_patterns = [
            r"C:\Python*\python.exe",
            r"C:\Python*\Scripts\python.exe",
        ]

        # 사용자 설치 경로
        user_home = Path.home()
        user_patterns = [
            str(user_home / "AppData" / "Local" / "Programs" / "Python" / "Python*" / "python.exe"),
            str(user_home / "AppData" / "Local" / "Programs" / "Python" / "Python*" / "Scripts" / "python.exe"),
        ]

        # Anaconda/Miniconda 경로
        conda_patterns = [
            str(user_home / "anaconda3" / "python.exe"),
            str(user_home / "miniconda3" / "python.exe"),
            r"C:\ProgramData\Anaconda3\python.exe",
            r"C:\ProgramData\Miniconda3\python.exe",
        ]

        all_patterns = python_patterns + user_patterns + conda_patterns

        for pattern in all_patterns:
            self._find_pythons_by_glob(pattern)

    def _detect_unix_paths(self):
        """Unix 계열 시스템에서 Python 찾기"""
        common_paths = [
            "/usr/bin/python3",
            "/usr/bin/python",
            "/usr/local/bin/python3",
            "/usr/local/bin/python",
            "/opt/python*/bin/python3",
            "/opt/python*/bin/python",
        ]

        # Homebrew 경로 (macOS)
        if sys.platform == "darwin":
            common_paths.extend([
                "/opt/homebrew/bin/python3",
                "/usr/local/bin/python3",
            ])

        # 사용자 홈 경로
        user_home = Path.home()
        user_paths = [
            str(user_home / ".pyenv" / "versions" / "*" / "bin" / "python"),
            str(user_home / "anaconda3" / "bin" / "python"),
            str(user_home / "miniconda3" / "bin" / "python"),
        ]

        all_paths = common_paths + user_paths

        for path_pattern in all_paths:
            if "*" in path_pattern:
                self._find_pythons_by_glob(path_pattern)
            else:
                if os.path.isfile(path_pattern) and os.access(path_pattern, os.X_OK):
                    info = self._get_python_info(path_pattern)
                    if info:
                        self.found_pythons.append(info)

    def _find_pythons_by_glob(self, pattern: str):
        """glob 패턴으로 Python 찾기"""
        try:
            import glob
            for path in glob.glob(pattern):
                if os.path.isfile(path) and os.access(path, os.X_OK):
                    info = self._get_python_info(path)
                    if info:
                        self.found_pythons.append(info)
        except Exception:
            pass  # glob 실패 시 무시

    def _get_python_info(self, python_path: str) -> Optional[PythonInfo]:
        """Python 경로에서 버전 정보 추출"""
        try:
            # python --version 실행
            result = subprocess.run(
                [python_path, "--version"],
                capture_output=True,
                text=True,
                timeout=5
            )

            if result.returncode == 0:
                version_output = result.stdout.strip()
                # "Python 3.13.0" 형태에서 버전 추출
                version_match = re.search(r"Python (\d+\.\d+\.\d+)", version_output)
                if version_match:
                    version = version_match.group(1)
                    return PythonInfo(python_path, version)

        except (subprocess.TimeoutExpired, subprocess.SubprocessError, FileNotFoundError):
            pass

        return None

    def _remove_duplicates(self):
        """중복된 Python 제거 (같은 경로 또는 같은 버전+경로)"""
        seen_paths = set()
        unique_pythons = []

        for python in self.found_pythons:
            # 경로 정규화
            normalized_path = os.path.normpath(python.path)
            if normalized_path not in seen_paths:
                seen_paths.add(normalized_path)
                python.path = normalized_path  # 정규화된 경로로 업데이트
                unique_pythons.append(python)

        self.found_pythons = unique_pythons

    def _sort_by_version(self):
        """버전별로 정렬 (높은 버전 우선)"""
        def version_key(python_info: PythonInfo) -> Tuple[int, int, int]:
            try:
                parts = python_info.version.split(".")
                return (int(parts[0]), int(parts[1]), int(parts[2]) if len(parts) > 2 else 0)
            except (ValueError, IndexError):
                return (0, 0, 0)

        self.found_pythons.sort(key=version_key, reverse=True)

    def _mark_recommended(self):
        """권장 버전 표시"""
        for python in self.found_pythons:
            try:
                parts = python.version.split(".")
                version_tuple = (int(parts[0]), int(parts[1]), int(parts[2]) if len(parts) > 2 else 0)
                if version_tuple >= self.RECOMMENDED_VERSION:
                    python.is_recommended = True
                    break  # 첫 번째 권장 버전만 표시
            except (ValueError, IndexError):
                continue


def detect_python_installations() -> List[PythonInfo]:
    """
    시스템에 설치된 Python을 탐지합니다.

    Returns:
        List[PythonInfo]: 발견된 Python 설치 목록
    """
    detector = PythonDetector()
    return detector.detect_all()


def get_best_python() -> Optional[PythonInfo]:
    """
    가장 적합한 Python 설치를 반환합니다.

    Returns:
        Optional[PythonInfo]: 가장 적합한 Python, 없으면 None
    """
    detector = PythonDetector()
    return detector.get_best_python()


def format_python_info_for_template(python_info: Optional[PythonInfo]) -> str:
    """
    템플릿에 삽입할 Python 정보를 포맷팅합니다.

    Args:
        python_info: Python 정보 객체

    Returns:
        str: 템플릿용 포맷팅된 문자열
    """
    if not python_info:
        return "Python이 설치되지 않았거나 탐지되지 않았습니다."

    return f"{python_info.path}"


if __name__ == "__main__":
    # 테스트 코드
    pythons = detect_python_installations()

    if pythons:
        print("발견된 Python 설치:")
        for i, python in enumerate(pythons, 1):
            marker = " (권장)" if python.is_recommended else ""
            print(f"{i}. {python}{marker}")

        best = get_best_python()
        if best:
            print(f"\n가장 적합한 Python: {best}")
    else:
        print("Python이 설치되지 않았거나 탐지되지 않았습니다.")