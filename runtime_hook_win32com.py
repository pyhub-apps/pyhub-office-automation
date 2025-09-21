#!/usr/bin/env python
"""
PyInstaller 런타임 훅: win32com 초기화 및 캐시 설정
Windows에서 PyInstaller로 빌드된 실행 파일의 COM 캐시 재구축 경고 방지
"""

import os
import sys
import tempfile
import warnings


def initialize_win32com():
    """PyInstaller 환경에서 win32com을 초기화합니다."""
    try:
        # PyInstaller 실행 환경 확인
        if hasattr(sys, "_MEIPASS"):
            # win32com import 전에 경고 억제
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")

                import win32com
                import win32com.client

                # 캐시 디렉터리 설정
                cache_dir = os.path.join(tempfile.gettempdir(), "win32com_gen_py")

                # 캐시 디렉터리 생성
                if not os.path.exists(cache_dir):
                    os.makedirs(cache_dir, exist_ok=True)

                # win32com 캐시 경로 설정
                win32com.__gen_path__ = cache_dir

                # gencache 설정
                win32com.client.gencache.is_readonly = False

                # 캐시 디렉터리 확인
                try:
                    win32com.client.gencache.GetGeneratePath()
                except:
                    # 캐시 생성 실패 시 읽기 전용으로 설정
                    win32com.client.gencache.is_readonly = True

    except ImportError:
        # win32com이 없는 환경에서는 무시
        pass
    except Exception:
        # 기타 오류 시 조용히 무시
        pass


# 런타임 훅 실행
if __name__ == "__main__":
    initialize_win32com()
else:
    # 모듈로 import될 때도 실행
    initialize_win32com()
