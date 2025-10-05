"""
Excel 엔진 예외 클래스 정의

플랫폼별 Excel 엔진에서 발생하는 예외들을 정의합니다.
"""


class ExcelEngineError(Exception):
    """Excel 엔진 기본 예외 클래스"""

    pass


class WorkbookNotFoundError(ExcelEngineError):
    """워크북을 찾을 수 없는 경우"""

    def __init__(self, workbook_name: str):
        self.workbook_name = workbook_name
        super().__init__(f"워크북 '{workbook_name}'을 찾을 수 없습니다")


class SheetNotFoundError(ExcelEngineError):
    """시트를 찾을 수 없는 경우"""

    def __init__(self, sheet_name: str):
        self.sheet_name = sheet_name
        super().__init__(f"시트 '{sheet_name}'을 찾을 수 없습니다")


class RangeError(ExcelEngineError):
    """범위 관련 오류"""

    def __init__(self, range_str: str, message: str):
        self.range_str = range_str
        super().__init__(f"범위 '{range_str}': {message}")


class TableNotFoundError(ExcelEngineError):
    """테이블을 찾을 수 없는 경우"""

    def __init__(self, table_name: str):
        self.table_name = table_name
        super().__init__(f"테이블 '{table_name}'을 찾을 수 없습니다")


class ChartNotFoundError(ExcelEngineError):
    """차트를 찾을 수 없는 경우"""

    def __init__(self, chart_name: str):
        self.chart_name = chart_name
        super().__init__(f"차트 '{chart_name}'을 찾을 수 없습니다")


class PlatformNotSupportedError(ExcelEngineError):
    """플랫폼이 지원되지 않는 경우"""

    def __init__(self, platform: str, feature: str = ""):
        self.platform = platform
        self.feature = feature
        if feature:
            message = f"'{feature}' 기능은 {platform}에서 지원되지 않습니다"
        else:
            message = f"{platform}은 지원되지 않는 플랫폼입니다. Windows 또는 macOS에서만 사용 가능합니다"
        super().__init__(message)


class ExcelNotRunningError(ExcelEngineError):
    """Excel이 실행되지 않은 경우"""

    def __init__(self):
        super().__init__("Excel이 실행되지 않았습니다. Excel을 먼저 실행하거나 워크북을 여세요")


class COMError(ExcelEngineError):
    """Windows COM 관련 오류"""

    def __init__(self, message: str):
        super().__init__(f"COM 오류: {message}")


class AppleScriptError(ExcelEngineError):
    """macOS AppleScript 관련 오류"""

    def __init__(self, script: str, stderr: str):
        self.script = script
        self.stderr = stderr
        super().__init__(f"AppleScript 실행 실패: {stderr}")


class DataValidationError(ExcelEngineError):
    """데이터 검증 오류"""

    def __init__(self, message: str, details: dict = None):
        self.details = details or {}
        super().__init__(f"데이터 검증 오류: {message}")


class EngineInitializationError(ExcelEngineError):
    """엔진 초기화 오류"""

    def __init__(self, engine_type: str, message: str):
        self.engine_type = engine_type
        super().__init__(f"{engine_type} 엔진 초기화 실패: {message}")
