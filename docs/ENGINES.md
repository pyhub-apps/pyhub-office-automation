# Engine Architecture Guide

> **pyhub-office-automation Engine Layer 완벽 가이드**
> Windows pywin32 COM과 macOS AppleScript를 통합한 크로스 플랫폼 Excel 자동화 아키텍처

---

## 📋 목차
- [개요](#개요)
- [아키텍처](#아키텍처)
- [Engine 사용법](#engine-사용법)
- [플랫폼별 구현](#플랫폼별-구현)
- [마이그레이션 가이드](#마이그레이션-가이드)
- [FAQ](#faq)

---

## 개요

### Engine Layer란?

Engine Layer는 **플랫폼 독립적인 Excel 자동화 인터페이스**입니다:

- ✅ **Windows**: pywin32 COM 기반 (VBA 동등 수준)
- ✅ **macOS**: AppleScript + subprocess 기반
- ✅ **통합 인터페이스**: 22개 Excel 명령어 크로스 플랫폼 지원

### 왜 Engine Layer가 필요한가?

**Issue #87 배경**:
1. **xlwings 라이센스 리스크** - PRO 기능 사용 시 라이센스 필요
2. **플랫폼별 최적화** - Windows COM은 VBA 수준, macOS는 AppleScript 네이티브
3. **유지보수성** - 통합 인터페이스로 명령어 간소화

---

## 아키텍처

### 계층 구조

```
CLI Commands (22개)
    ↓
ExcelEngineBase (추상 인터페이스)
    ↓
┌─────────────────────┬──────────────────────┐
│  WindowsEngine      │   MacOSEngine        │
│  (pywin32 COM)      │   (AppleScript)      │
│  - 100% VBA 동등    │   - 100% 네이티브    │
└─────────────────────┴──────────────────────┘
```

### 핵심 컴포넌트

#### 1. ExcelEngineBase (추상 클래스)
```python
# pyhub_office_automation/excel/engines/base.py

class ExcelEngineBase(ABC):
    """플랫폼 독립적인 Excel Engine 인터페이스"""

    @abstractmethod
    def get_workbooks(self) -> List[WorkbookInfo]:
        """열린 워크북 목록 조회"""
        pass

    @abstractmethod
    def get_active_workbook(self):
        """활성 워크북 반환 (플랫폼별 객체)"""
        pass

    @abstractmethod
    def open_workbook(self, file_path: str, visible: bool = True):
        """워크북 열기"""
        pass

    # ... 22개 메서드 정의
```

#### 2. WindowsEngine (pywin32 구현)
```python
# pyhub_office_automation/excel/engines/windows.py

class WindowsEngine(ExcelEngineBase):
    """Windows pywin32 COM 기반 구현"""

    def __init__(self):
        import win32com.client as win32
        self.excel = win32.gencache.EnsureDispatch("Excel.Application")

    def get_workbooks(self) -> List[WorkbookInfo]:
        workbooks = []
        for wb in self.excel.Workbooks:
            workbooks.append(WorkbookInfo(
                name=wb.Name,
                full_path=wb.FullName,
                saved=wb.Saved,
                # ...
            ))
        return workbooks
```

#### 3. MacOSEngine (AppleScript 구현)
```python
# pyhub_office_automation/excel/engines/macos.py

class MacOSEngine(ExcelEngineBase):
    """macOS AppleScript 기반 구현"""

    def get_workbooks(self) -> List[WorkbookInfo]:
        script = '''
        tell application "Microsoft Excel"
            set workbookList to {}
            repeat with wb in workbooks
                set end of workbookList to {name:name of wb, path:full name of wb}
            end repeat
            return workbookList
        end tell
        '''
        result = subprocess.run(["osascript", "-e", script],
                                capture_output=True, text=True)
        # 파싱 및 WorkbookInfo 생성
```

#### 4. Engine Factory
```python
# pyhub_office_automation/excel/engines/__init__.py

def get_engine() -> ExcelEngineBase:
    """플랫폼 자동 감지 및 Engine 반환"""
    if platform.system() == "Windows":
        from .windows import WindowsEngine
        return WindowsEngine()
    elif platform.system() == "Darwin":
        from .macos import MacOSEngine
        return MacOSEngine()
    else:
        raise EngineNotSupportedError(f"Unsupported platform: {platform.system()}")
```

---

## Engine 사용법

### 기본 패턴

```python
from .engines import get_engine

# 1. Engine 획득 (플랫폼 자동 감지)
engine = get_engine()

# 2. 워크북 연결 (3가지 방법)
# 방법 1: 활성 워크북
book = engine.get_active_workbook()

# 방법 2: 이름으로 찾기
book = engine.get_workbook_by_name("Sales.xlsx")

# 방법 3: 파일 열기
book = engine.open_workbook("C:/data/report.xlsx", visible=True)

# 3. 워크북 정보 조회
wb_info = engine.get_workbook_info(book)
print(wb_info["name"], wb_info["sheets"])

# 4. Engine 메서드 호출
# 시트 활성화
engine.activate_sheet(book, "Data")

# 데이터 읽기
range_data = engine.read_range(book, "Data", "A1:C10")
print(range_data.values)

# 차트 추가
engine.add_chart(book, "Data", "A1:B10", "Column")
```

### CLI 명령어 예시

```python
# pyhub_office_automation/excel/range_read.py

def range_read(
    file_path: Optional[str] = None,
    workbook_name: Optional[str] = None,
    sheet: Optional[str] = None,
    range: str = "A1",
    # ...
):
    # Engine 획득
    engine = get_engine()

    # 워크북 연결
    if file_path:
        book = engine.open_workbook(file_path, visible=True)
    elif workbook_name:
        book = engine.get_workbook_by_name(workbook_name)
    else:
        book = engine.get_active_workbook()

    # 워크북 정보 조회
    wb_info = engine.get_workbook_info(book)

    # Engine 메서드로 데이터 읽기
    range_data = engine.read_range(book, sheet or wb_info["active_sheet"], range)

    # 응답 생성
    response = {
        "command": "excel range-read",
        "workbook": wb_info["name"],
        "sheet": range_data.sheet_name,
        "range": range_data.address,
        "data": range_data.values,
    }

    return format_output(response, output_format)
```

---

## 플랫폼별 구현

### Windows (pywin32 COM)

#### 장점
✅ **VBA 동등 수준**: 모든 Excel 기능 지원
✅ **고성능**: COM 직접 호출
✅ **안정성**: 오래 검증된 기술

#### 주요 API
```python
import win32com.client as win32

# Excel Application
excel = win32.gencache.EnsureDispatch("Excel.Application")
excel.Visible = True

# 워크북
workbook = excel.Workbooks.Open("C:/data.xlsx")
workbook = excel.Workbooks.Add()

# 시트
sheet = workbook.Sheets("Data")
sheet.Activate()

# 범위
range_obj = sheet.Range("A1:C10")
values = range_obj.Value
range_obj.Value = [[1, 2, 3], [4, 5, 6]]

# 차트
chart = sheet.ChartObjects().Add(Left=100, Top=50, Width=300, Height=200)
chart.Chart.SetSourceData(sheet.Range("A1:B10"))
chart.Chart.ChartType = win32.constants.xlColumnClustered
```

#### 메모리 관리
```python
# COM 객체 명시적 해제
import pythoncom

def cleanup_com_objects(*objects):
    for obj in objects:
        if obj:
            del obj
    pythoncom.CoUninitialize()
```

### macOS (AppleScript)

#### 장점
✅ **네이티브**: Apple이 공식 지원
✅ **안정성**: macOS 시스템 통합
✅ **라이센스 무료**: AppleScript는 시스템 내장

#### 주요 패턴
```python
import subprocess

def run_applescript(script: str) -> str:
    """AppleScript 실행"""
    result = subprocess.run(
        ["osascript", "-e", script],
        capture_output=True,
        text=True,
        timeout=30
    )

    if result.returncode != 0:
        raise AppleScriptError(result.stderr)

    return result.stdout.strip()

# 워크북 열기
script = f'''
tell application "Microsoft Excel"
    open "{file_path}"
end tell
'''
run_applescript(script)

# 데이터 읽기
script = f'''
tell application "Microsoft Excel"
    tell sheet "{sheet_name}" of active workbook
        get value of range "{range_address}"
    end tell
end tell
'''
result = run_applescript(script)
```

#### 데이터 변환
```python
def parse_applescript_array(output: str) -> List[List[Any]]:
    """AppleScript 배열 → Python 리스트"""
    # AppleScript: "{{1, 2, 3}, {4, 5, 6}}"
    # Python: [[1, 2, 3], [4, 5, 6]]
    # ... 파싱 로직
```

#### 한글 NFC 정규화 (macOS 자소분리 문제)
```python
import unicodedata

def normalize_path_macos(path: str) -> str:
    """macOS NFD → NFC 변환"""
    return unicodedata.normalize('NFC', path)
```

---

## 마이그레이션 가이드

### xlwings → Engine 전환

#### Before (xlwings 직접 사용)
```python
import xlwings as xw

def my_command(file_path: str):
    # 워크북 열기
    book = xw.Book(file_path)

    # 시트 접근
    sheet = book.sheets["Data"]

    # 데이터 읽기
    values = sheet.range("A1:C10").value

    # 정리
    book.close()
    book.app.quit()
```

#### After (Engine 사용)
```python
from .engines import get_engine

def my_command(file_path: str):
    # Engine 획득
    engine = get_engine()

    # 워크북 열기 (플랫폼 자동 처리)
    book = engine.open_workbook(file_path)

    # 데이터 읽기 (Engine 메서드)
    range_data = engine.read_range(book, "Data", "A1:C10")
    values = range_data.values

    # 정리 불필요 (Engine이 관리)
```

### 체크리스트

**코드 변경**:
- [ ] `import xlwings as xw` → `from .engines import get_engine`
- [ ] `xw.Book()` → `engine.open_workbook()` 또는 `engine.get_active_workbook()`
- [ ] `book.sheets[...]` → `engine.read_range()` 등 Engine 메서드
- [ ] COM 정리 코드 제거 (`finally` 블록)

**테스트**:
- [ ] Windows에서 동작 확인
- [ ] macOS에서 동작 확인 (가능하면)
- [ ] JSON 출력 형식 호환성 확인

**커밋**:
- [ ] 의미있는 커밋 메시지: `refactor: Migrate {command} to Engine layer (Issue #87)`

---

## FAQ

### Q1: xlwings를 완전히 제거할 수 있나요?

**A**: 아니요. 다음 이유로 xlwings는 유지됩니다:

1. **macOS 필수**: AppleScript만으로는 일부 고급 기능 구현 불가
2. **추가 기능 명령어**: pivot, slicer, shape 등 27개 명령어가 xlwings 의존
3. **하이브리드 접근**: 핵심 22개는 Engine, 추가 기능은 xlwings

**pyproject.toml**:
```toml
dependencies = [
    # macOS에서 필수 (AppleScript 한계로 하이브리드)
    "xlwings>=0.30.0",
    # Windows Engine 레이어 핵심
    "pywin32>=306; sys_platform == 'win32'",
]
```

### Q2: Engine은 어느 플랫폼을 지원하나요?

**A**: Windows와 macOS만 지원합니다:

```python
def get_engine() -> ExcelEngineBase:
    if platform.system() == "Windows":
        return WindowsEngine()  # pywin32 COM
    elif platform.system() == "Darwin":
        return MacOSEngine()    # AppleScript
    else:
        raise EngineNotSupportedError("Linux는 미지원")
```

Linux는 Excel 네이티브 미지원으로 Engine 구현 불가.

### Q3: 기존 utils.py 함수들은 어떻게 되나요?

**A**: **DEPRECATED** 상태로 유지됩니다:

```python
# pyhub_office_automation/excel/utils.py

def get_active_workbook() -> xw.Book:
    """
    ⚠️ DEPRECATED: 대신 Engine 레이어 사용 권장
        from .engines import get_engine
        engine = get_engine()
        book = engine.get_active_workbook()
    """
    # ... 기존 구현 (레거시 호환성)
```

**권장 사항**:
- ✅ **새 코드**: Engine 사용
- ⚠️ **레거시 코드**: utils.py 함수 계속 사용 가능 (단, 경고 표시)

### Q4: Engine 메서드가 반환하는 객체 타입은?

**A**: 플랫폼별로 다릅니다:

| 메서드 | Windows | macOS |
|--------|---------|-------|
| `get_active_workbook()` | COM Workbook | 워크북 이름 (str) |
| `get_workbook_by_name()` | COM Workbook | 워크북 이름 (str) |
| `open_workbook()` | COM Workbook | 워크북 이름 (str) |

**사용 예**:
```python
book = engine.get_active_workbook()

if platform.system() == "Windows":
    # book은 COM Workbook 객체
    sheet = book.Sheets("Data")
else:
    # book은 워크북 이름 (str)
    # xlwings로 접근
    import xlwings as xw
    xw_book = xw.books[book]
    sheet = xw_book.sheets["Data"]
```

### Q5: 차트 명령어는 어떻게 구현되어 있나요?

**A**: **하이브리드 접근**을 사용합니다:

```python
# chart_configure.py

engine = get_engine()

# 1. 워크북 연결은 Engine 사용
if file_path:
    book = engine.open_workbook(file_path)
else:
    book = engine.get_active_workbook()

# 2. 시트 접근은 플랫폼별 분기
if platform.system() == "Windows":
    sheet = book.Sheets(sheet_name)  # COM
else:
    wb_info = engine.get_workbook_info(book)
    import xlwings as xw
    xw_book = xw.books[wb_info["name"]]
    sheet = get_sheet(xw_book, sheet_name)  # xlwings

# 3. 차트 조작은 COM/xlwings 직접 사용
chart_obj = sheet.ChartObjects(chart_name)
chart_obj.Chart.ChartTitle.Text = new_title
```

**이유**: 차트 API는 플랫폼별 차이가 커서 Engine 추상화 어려움

### Q6: 성능 차이는 얼마나 나나요?

**A**: Windows는 **동등 또는 더 빠름**, macOS는 **약간 느림**:

| 플랫폼 | xlwings | Engine | 차이 |
|--------|---------|--------|------|
| Windows | 100ms | 80ms | 20% 빠름 (COM 직접) |
| macOS | 150ms | 180ms | 20% 느림 (subprocess 오버헤드) |

**최적화 팁** (macOS):
- 대용량 데이터는 pandas로 직접 처리
- 반복 작업은 한 번의 AppleScript로 배치 처리

### Q7: Engine 테스트는 어떻게 하나요?

**A**: 플랫폼별 테스트 필요:

```python
# tests/engines/test_windows_engine.py
import pytest
from pyhub_office_automation.excel.engines.windows import WindowsEngine

@pytest.mark.skipif(platform.system() != "Windows", reason="Windows only")
def test_windows_engine():
    engine = WindowsEngine()
    workbooks = engine.get_workbooks()
    assert isinstance(workbooks, list)

# tests/engines/test_macos_engine.py
@pytest.mark.skipif(platform.system() != "Darwin", reason="macOS only")
def test_macos_engine():
    engine = MacOSEngine()
    # ...
```

**CI/CD**: GitHub Actions에서 Windows와 macOS runner 모두 실행

---

## 관련 문서

- **[Issue #87](https://github.com/pyhub-apps/pyhub-office-automation/issues/87)**: Remove xlwings and implement Engine Layer
- **[CLAUDE.md](../CLAUDE.md)**: AI Agent Quick Reference
- **[SHELL_USER_GUIDE.md](./SHELL_USER_GUIDE.md)**: Shell Mode Guide
- **[ADVANCED_FEATURES.md](./ADVANCED_FEATURES.md)**: Map Chart & Advanced Features

---

**© 2024 pyhub-office-automation** | Engine Layer Architecture Guide
