"""
pytest 설정 및 공통 fixture
"""

import gc
import platform
import sys
import tempfile
from pathlib import Path
from unittest.mock import Mock, patch

import pytest

# 프로젝트 루트를 Python 경로에 추가
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))


@pytest.fixture(scope="session", autouse=True)
def cleanup_excel_after_tests():
    """
    테스트 세션 종료 후 남은 Excel 인스턴스 정리
    """
    yield

    # 테스트 완료 후 Excel 인스턴스 정리
    try:
        import xlwings as xw

        # 모든 열려있는 Excel 앱 종료
        try:
            # 현재 활성 앱이 있으면 종료
            apps = xw.apps
            for app in apps:
                try:
                    app.quit()
                except:
                    pass
        except:
            pass

        # 가비지 컬렉션
        for _ in range(3):
            gc.collect()

        # Windows에서 COM 정리
        if platform.system() == "Windows":
            try:
                import pythoncom

                pythoncom.CoUninitialize()
            except:
                pass

    except Exception as e:
        print(f"Excel cleanup failed: {e}")


@pytest.fixture
def temp_excel_file():
    """임시 Excel 파일 생성 (테스트용)"""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as temp_file:
        temp_path = Path(temp_file.name)

    yield temp_path

    # 정리
    if temp_path.exists():
        temp_path.unlink()


@pytest.fixture
def temp_invalid_file():
    """임시 텍스트 파일 생성 (잘못된 확장자 테스트용)"""
    with tempfile.NamedTemporaryFile(suffix=".txt", delete=False) as temp_file:
        temp_file.write(b"This is not an Excel file")
        temp_path = Path(temp_file.name)

    yield temp_path

    # 정리
    if temp_path.exists():
        temp_path.unlink()


@pytest.fixture
def mock_xlwings():
    """xlwings 모듈 모킹"""
    with patch("xlwings.App") as mock_app_class:
        # Mock App 인스턴스 설정
        mock_app = Mock()
        mock_app.visible = True
        mock_app_class.return_value = mock_app

        # Mock 워크북 설정
        mock_book = Mock()
        mock_book.name = "test_workbook.xlsx"
        mock_book.fullname = "/path/to/test_workbook.xlsx"
        mock_book.saved = True

        # Mock 시트 설정
        mock_sheet = Mock()
        mock_sheet.name = "Sheet1"
        mock_sheet.index = 1
        mock_sheet.visible = True

        # Mock 사용된 범위 설정
        mock_used_range = Mock()
        mock_used_range.last_cell.address = "C5"
        mock_used_range.rows.count = 5
        mock_used_range.columns.count = 3
        mock_sheet.used_range = mock_used_range

        # Mock sheets 컬렉션 설정
        mock_sheets = Mock()
        mock_sheets.__iter__ = lambda self: iter([mock_sheet])
        mock_sheets.__len__ = lambda self: 1
        mock_sheets.active = mock_sheet
        mock_book.sheets = mock_sheets

        mock_app.books.open.return_value = mock_book

        yield {
            "app_class": mock_app_class,
            "app": mock_app,
            "book": mock_book,
            "sheet": mock_sheet,
            "used_range": mock_used_range,
        }


@pytest.fixture
def mock_xlwings_error():
    """xlwings 에러 상황 모킹"""
    with patch("xlwings.App") as mock_app_class:
        mock_app_class.side_effect = Exception("Excel을 시작할 수 없습니다")
        yield mock_app_class


@pytest.fixture
def non_existent_file():
    """존재하지 않는 파일 경로"""
    return "/path/to/non_existent_file.xlsx"
