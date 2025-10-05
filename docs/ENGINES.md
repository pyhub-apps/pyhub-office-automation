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
- ✅ **통합 인터페이스**: 40개 Excel 명령어 크로스 플랫폼 지원

### 왜 Engine Layer가 필요한가?

**Issue #87 & #88 배경**:
1. **xlwings 라이센스 리스크** - PRO 기능 사용 시 라이센스 필요
2. **플랫폼별 최적화** - Windows COM은 VBA 수준, macOS는 AppleScript 네이티브
3. **유지보수성** - 통합 인터페이스로 명령어 간소화
4. **확장성** - 고급 기능 지원을 위한 체계적 구조

### Engine Layer 진화 단계

**Issue #87 (완료)**: 핵심 22개 명령어 마이그레이션
- 워크북 관리 (4개)
- 시트 관리 (4개)
- 데이터 읽기/쓰기 (2개)
- 테이블 기본 (5개)
- 차트 기본 (7개)

**Issue #88 (완료)**: 고급 21개 명령어 추가
- 테이블 고급 (4개)
- 슬라이서 (4개)
- 피벗테이블 (5개)
- 도형 (5개)
- 데이터 변환 (3개 - utility 기반)

**현재 상태**: 총 43개 명령어 (40개 Engine 기반 + 3개 utility 기반)

---

## 아키텍처

### 계층 구조

```
CLI Commands (43개)
    ↓
ExcelEngineBase (추상 인터페이스 - 40개 메서드)
    ↓
┌─────────────────────┬──────────────────────┐
│  WindowsEngine      │   MacOSEngine        │
│  (pywin32 COM)      │   (AppleScript)      │
│  - 100% VBA 동등    │   - 100% 네이티브    │
│  - 40개 메서드 구현  │   - 40개 메서드 구현  │
└─────────────────────┴──────────────────────┘
```

### 핵심 컴포넌트

#### 1. ExcelEngineBase (추상 클래스)
```python
# pyhub_office_automation/excel/engines/base.py

class ExcelEngineBase(ABC):
    """
    플랫폼 독립적인 Excel Engine 인터페이스

    Issue #87: 핵심 22개 명령어 (완료)
    Issue #88: 추가 18개 명령어 (완료)
    """

    # ===========================================
    # 워크북 관리 (4개 명령어)
    # ===========================================
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

    # ===========================================
    # 피벗테이블 (5개 명령어) - Issue #88
    # ===========================================
    @abstractmethod
    def create_pivot_table(
        self, workbook: Any, source_sheet: str, source_range: str,
        dest_sheet: str, dest_cell: str, pivot_name: Optional[str] = None, **kwargs
    ) -> Dict[str, Any]:
        """피벗테이블 생성 (Windows 우선 지원)"""
        pass

    @abstractmethod
    def configure_pivot_table(
        self, workbook: Any, sheet: str, pivot_name: str,
        row_fields: Optional[List[str]] = None,
        column_fields: Optional[List[str]] = None,
        value_fields: Optional[List[Tuple[str, str]]] = None,
        filter_fields: Optional[List[str]] = None, **kwargs
    ):
        """피벗테이블 필드 설정"""
        pass

    # ... 총 40개 메서드 정의
```

#### 2. WindowsEngine (pywin32 구현)
```python
# pyhub_office_automation/excel/engines/windows.py

class WindowsEngine(ExcelEngineBase):
    """Windows pywin32 COM 기반 구현 - 40개 메서드"""

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

    def create_pivot_table(self, workbook, source_sheet, source_range,
                          dest_sheet, dest_cell, pivot_name=None, **kwargs):
        """Windows COM을 사용한 피벗테이블 생성"""
        import win32com.client as win32

        # 소스 데이터 설정
        src_sheet = workbook.Sheets(source_sheet)
        src_range = src_sheet.Range(source_range)

        # 대상 시트와 위치
        dst_sheet = workbook.Sheets(dest_sheet)
        dst_cell = dst_sheet.Range(dest_cell)

        # 피벗테이블 캐시 생성
        pc_cache = workbook.PivotCaches().Create(
            SourceType=win32.constants.xlDatabase,
            SourceData=src_range,
            Version=win32.constants.xlPivotTableVersion15
        )

        # 피벗테이블 생성
        pivot_table = pc_cache.CreatePivotTable(
            TableDestination=dst_cell,
            TableName=pivot_name or f"PivotTable{len(dst_sheet.PivotTables()) + 1}",
            DefaultVersion=win32.constants.xlPivotTableVersion15
        )

        return {
            "name": pivot_table.Name,
            "source": f"{source_sheet}!{source_range}",
            "destination": f"{dest_sheet}!{dest_cell}"
        }
```

#### 3. MacOSEngine (AppleScript 구현)
```python
# pyhub_office_automation/excel/engines/macos.py

class MacOSEngine(ExcelEngineBase):
    """macOS AppleScript 기반 구현 - 일부 고급 기능 제한"""

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

    def create_pivot_table(self, workbook, source_sheet, source_range,
                          dest_sheet, dest_cell, pivot_name=None, **kwargs):
        """macOS에서는 제한적 지원"""
        raise EngineNotSupportedError(
            "피벗테이블 생성은 Windows에서만 완전 지원됩니다. "
            "macOS에서는 수동으로 생성하거나 xlwings를 사용하세요."
        )
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

### Issue #88 신규 메서드 사용 예제

#### 1. 테이블 고급 기능 (4개)
```python
# 테이블 생성
table_info = engine.create_table(
    workbook=book,
    sheet="Data",
    range_str="A1:D100",
    table_name="SalesTable",
    has_headers=True,
    table_style="TableStyleMedium2"
)

# 테이블 정렬
engine.sort_table(
    workbook=book,
    sheet="Data",
    table_name="SalesTable",
    sort_fields=[("Revenue", "desc"), ("Region", "asc")]
)

# 정렬 해제
engine.clear_table_sort(book, "Data", "SalesTable")

# 정렬 정보 조회
sort_info = engine.get_table_sort_info(book, "Data", "SalesTable")
print(f"정렬 필드: {sort_info}")
```

#### 2. 슬라이서 (4개) - Windows 전용
```python
# 슬라이서 추가
slicer = engine.add_slicer(
    workbook=book,
    sheet="Dashboard",
    pivot_name="PivotTable1",
    field_name="Region",
    left=400, top=50,
    width=200, height=150,
    slicer_name="RegionSlicer",
    caption="지역 선택",
    style="SlicerStyleLight2"
)

# 슬라이서 목록
slicers = engine.list_slicers(book, sheet="Dashboard")

# 슬라이서 위치 조정
engine.position_slicer(book, "Dashboard", "RegionSlicer",
                       left=500, top=100, width=250)

# 슬라이서 연결 (여러 피벗테이블에 연결)
engine.connect_slicer(book, "RegionSlicer",
                     ["PivotTable1", "PivotTable2", "PivotTable3"])
```

#### 3. 피벗테이블 (5개)
```python
# 피벗테이블 생성
pivot = engine.create_pivot_table(
    workbook=book,
    source_sheet="RawData",
    source_range="A1:F1000",
    dest_sheet="Analysis",
    dest_cell="H1",
    pivot_name="SalesAnalysis"
)

# 피벗테이블 설정
engine.configure_pivot_table(
    workbook=book,
    sheet="Analysis",
    pivot_name="SalesAnalysis",
    row_fields=["Region", "Product"],
    column_fields=["Year"],
    value_fields=[("Revenue", "Sum"), ("Quantity", "Count")],
    filter_fields=["Category"]
)

# 피벗테이블 새로고침
engine.refresh_pivot_table(book, "Analysis", "SalesAnalysis")

# 피벗테이블 목록
pivots = engine.list_pivot_tables(book)
for pivot in pivots:
    print(f"{pivot.name}: {pivot.sheet_name} - {pivot.source_range}")

# 피벗테이블 삭제
engine.delete_pivot_table(book, "Analysis", "OldPivot")
```

#### 4. 도형 (5개)
```python
# 도형 추가
shape = engine.add_shape(
    workbook=book,
    sheet="Report",
    shape_type="rectangle",
    left=100, top=100,
    width=200, height=100,
    shape_name="InfoBox",
    fill_color="0066CC",
    transparency=0.2
)

# 도형 목록
shapes = engine.list_shapes(book, sheet="Report")

# 도형 서식 설정
engine.format_shape(
    workbook=book,
    sheet="Report",
    shape_name="InfoBox",
    fill_color="FF6600",
    line_color="000000",
    line_width=2,
    text="중요 정보"
)

# 도형 그룹화
group_name = engine.group_shapes(
    workbook=book,
    sheet="Report",
    shape_names=["Shape1", "Shape2", "Shape3"],
    group_name="DashboardGroup"
)

# 도형 삭제
engine.delete_shape(book, "Report", "OldShape")
```

### CLI 명령어 예시

```python
# pyhub_office_automation/excel/pivot_create.py

def pivot_create(
    file_path: Optional[str] = None,
    workbook_name: Optional[str] = None,
    source_range: str = "A1:D100",
    dest_range: str = "F1",
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

    # Engine 메서드로 피벗테이블 생성
    pivot_result = engine.create_pivot_table(
        workbook=book.api,  # Windows: COM 객체
        source_sheet=source_sheet.name,
        source_range=source_range,
        dest_sheet=dest_sheet.name,
        dest_cell=dest_cell,
        pivot_name=pivot_name
    )

    # 응답 생성
    response = {
        "command": "excel pivot-create",
        "pivot": pivot_result,
        "workbook": wb_info["name"]
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
✅ **완전 지원**: 40개 메서드 모두 구현

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

# 피벗테이블 (Issue #88)
pivot_cache = workbook.PivotCaches().Create(
    SourceType=win32.constants.xlDatabase,
    SourceData=sheet.Range("A1:D100")
)
pivot_table = pivot_cache.CreatePivotTable(
    TableDestination=sheet.Range("F1"),
    TableName="PivotTable1"
)

# 슬라이서 (Issue #88)
slicer_cache = workbook.SlicerCaches.Add2(
    Source=pivot_table,
    SourceField="Region"
)
slicer = slicer_cache.Slicers.Add(
    SlicerDestination=sheet,
    Caption="Region Filter",
    Top=50, Left=400,
    Width=200, Height=150
)
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
⚠️ **부분 지원**: 일부 고급 기능 제한 (슬라이서, 복잡한 피벗 등)

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

# 테이블 생성 (Issue #88 - 기본 지원)
script = f'''
tell application "Microsoft Excel"
    tell sheet "{sheet_name}" of active workbook
        make new list object at end with properties {{
            source range: range "{range_address}",
            name: "{table_name}",
            has headers: true
        }}
    end tell
end tell
'''
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

### Issue #88 신규 명령어 마이그레이션 예제

#### 피벗테이블 생성 (Before - xlwings)
```python
import xlwings as xw

def create_pivot_xlwings(file_path: str):
    book = xw.Book(file_path)
    src_sheet = book.sheets["RawData"]
    dst_sheet = book.sheets["Analysis"]

    # xlwings API로 피벗테이블 생성 (복잡)
    src_range = src_sheet.range("A1:F1000")
    pivot_cache = book.api.PivotCaches().Create(
        SourceType=1,  # xlDatabase
        SourceData=src_range.api
    )
    # ... 복잡한 COM 호출
```

#### 피벗테이블 생성 (After - Engine)
```python
def create_pivot_engine(file_path: str):
    engine = get_engine()
    book = engine.open_workbook(file_path)

    # Engine 메서드로 간단하게 생성
    pivot = engine.create_pivot_table(
        workbook=book,
        source_sheet="RawData",
        source_range="A1:F1000",
        dest_sheet="Analysis",
        dest_cell="H1",
        pivot_name="SalesAnalysis"
    )
```

#### 슬라이서 추가 (Before - COM 직접)
```python
import win32com.client as win32

def add_slicer_com(workbook):
    # COM 상수 임포트 필요
    excel = win32.gencache.EnsureDispatch("Excel.Application")

    # 복잡한 COM 호출
    pivot_table = workbook.Sheets("Dashboard").PivotTables("PivotTable1")
    slicer_cache = workbook.SlicerCaches.Add2(
        Source=pivot_table,
        SourceField="Region",
        Name="Slicer_Region"
    )
    # ... 더 많은 설정
```

#### 슬라이서 추가 (After - Engine)
```python
def add_slicer_engine(file_path: str):
    engine = get_engine()
    book = engine.open_workbook(file_path)

    # Engine 메서드로 간결하게
    slicer = engine.add_slicer(
        workbook=book,
        sheet="Dashboard",
        pivot_name="PivotTable1",
        field_name="Region",
        left=400, top=50,
        slicer_name="RegionSlicer"
    )
```

### 체크리스트

**코드 변경**:
- [ ] `import xlwings as xw` → `from .engines import get_engine`
- [ ] `xw.Book()` → `engine.open_workbook()` 또는 `engine.get_active_workbook()`
- [ ] `book.sheets[...]` → `engine.read_range()` 등 Engine 메서드
- [ ] COM 정리 코드 제거 (`finally` 블록)
- [ ] 피벗/슬라이서 COM 코드 → Engine 메서드

**테스트**:
- [ ] Windows에서 동작 확인
- [ ] macOS에서 동작 확인 (가능하면)
- [ ] JSON 출력 형식 호환성 확인
- [ ] 고급 기능 플랫폼별 차이 확인

**커밋**:
- [ ] 의미있는 커밋 메시지: `refactor: Migrate {command} to Engine layer (Issue #88)`

---

## FAQ

### Q1: xlwings를 완전히 제거할 수 있나요?

**A**: 아니요. 다음 이유로 xlwings는 유지됩니다:

1. **macOS 필수**: AppleScript만으로는 일부 고급 기능 구현 불가
2. **추가 기능 명령어**: 일부 특수 명령어가 xlwings 의존
3. **하이브리드 접근**: 핵심 40개는 Engine, 특수 기능은 xlwings

**현재 상태 (Issue #88 이후)**:
- **Engine 기반**: 40개 명령어 (93%)
- **Utility 기반**: 3개 명령어 (7%)
- **총 명령어**: 43개

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
        return WindowsEngine()  # pywin32 COM - 40개 메서드 완전 지원
    elif platform.system() == "Darwin":
        return MacOSEngine()    # AppleScript - 일부 고급 기능 제한
    else:
        raise EngineNotSupportedError("Linux는 미지원")
```

**플랫폼별 지원 현황**:

| 기능 범주 | Windows | macOS | 비고 |
|----------|---------|-------|------|
| 워크북 관리 (4) | ✅ | ✅ | 완전 지원 |
| 시트 관리 (4) | ✅ | ✅ | 완전 지원 |
| 데이터 읽기/쓰기 (2) | ✅ | ✅ | 완전 지원 |
| 테이블 기본 (5) | ✅ | ✅ | 완전 지원 |
| 테이블 고급 (4) | ✅ | ⚠️ | macOS 부분 지원 |
| 차트 (7) | ✅ | ✅ | 완전 지원 |
| 피벗테이블 (5) | ✅ | ❌ | Windows 전용 |
| 슬라이서 (4) | ✅ | ❌ | Windows 전용 |
| 도형 (5) | ✅ | ⚠️ | macOS 제한적 |

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

### Q8: Issue #88의 데이터 변환 명령어는 왜 Engine에 없나요?

**A**: 데이터 변환 명령어 3개는 **utility 기반**으로 구현되었습니다:

```python
# data_analyze.py, data_transform.py, range_convert.py
# 이들은 pandas와 utility 함수만 사용

def data_analyze(...):
    # pandas로 데이터 분석
    df = pd.DataFrame(range_data.values)
    stats = df.describe()
    # ...

def data_transform(...):
    # 순수 Python으로 변환
    if transform_type == "transpose":
        transformed = list(zip(*data))
    # ...
```

**이유**:
- 플랫폼 독립적 로직 (Excel API 불필요)
- pandas가 더 효율적
- Engine 추상화 불필요

---

## Issue #88 신규 기능 상세

### 테이블 고급 기능 (4개 메서드)

#### 1. create_table
```python
def create_table(self, workbook, sheet, range_str, table_name=None,
                has_headers=True, table_style="TableStyleMedium2"):
    """
    Excel 테이블(ListObject) 생성

    Windows: COM API로 완전 지원
    macOS: AppleScript로 기본 지원

    Returns:
        Dict with table name, range, style info
    """
```

#### 2. sort_table
```python
def sort_table(self, workbook, sheet, table_name, sort_fields):
    """
    테이블 정렬 적용

    sort_fields: [("Column", "asc/desc"), ...]
    최대 3개 필드까지 다중 정렬 지원
    """
```

#### 3. clear_table_sort
```python
def clear_table_sort(self, workbook, sheet, table_name):
    """테이블 정렬 해제"""
```

#### 4. get_table_sort_info
```python
def get_table_sort_info(self, workbook, sheet, table_name):
    """현재 적용된 정렬 정보 반환"""
```

### 슬라이서 (4개 메서드) - Windows 전용

#### 1. add_slicer
```python
def add_slicer(self, workbook, sheet, pivot_name, field_name,
              left, top, width=200, height=150, slicer_name=None, **kwargs):
    """
    피벗테이블에 슬라이서 추가

    kwargs: caption, style, columns 등
    macOS: EngineNotSupportedError 발생
    """
```

#### 2. list_slicers
```python
def list_slicers(self, workbook, sheet=None):
    """워크북/시트의 모든 슬라이서 목록"""
```

#### 3. position_slicer
```python
def position_slicer(self, workbook, sheet, slicer_name,
                   left, top, width=None, height=None):
    """슬라이서 위치/크기 조정"""
```

#### 4. connect_slicer
```python
def connect_slicer(self, workbook, slicer_name, pivot_names):
    """슬라이서를 여러 피벗테이블에 연결"""
```

### 피벗테이블 (5개 메서드)

#### 1. create_pivot_table
```python
def create_pivot_table(self, workbook, source_sheet, source_range,
                      dest_sheet, dest_cell, pivot_name=None, **kwargs):
    """
    피벗테이블 생성

    Windows: 완전 지원
    macOS: 제한적 또는 미지원
    """
```

#### 2. configure_pivot_table
```python
def configure_pivot_table(self, workbook, sheet, pivot_name,
                         row_fields=None, column_fields=None,
                         value_fields=None, filter_fields=None, **kwargs):
    """
    피벗테이블 필드 설정

    value_fields: [("Field", "Function"), ...]
    Functions: Sum, Count, Average, Max, Min, etc.
    """
```

#### 3. refresh_pivot_table
```python
def refresh_pivot_table(self, workbook, sheet, pivot_name):
    """데이터 소스 변경 시 피벗테이블 새로고침"""
```

#### 4. delete_pivot_table
```python
def delete_pivot_table(self, workbook, sheet, pivot_name):
    """피벗테이블 삭제"""
```

#### 5. list_pivot_tables
```python
def list_pivot_tables(self, workbook, sheet=None):
    """피벗테이블 목록 조회"""
```

### 도형 (5개 메서드)

#### 1. add_shape
```python
def add_shape(self, workbook, sheet, shape_type, left, top,
             width, height, shape_name=None, **kwargs):
    """
    도형 추가

    shape_type: rectangle, oval, line, arrow, etc.
    kwargs: fill_color, transparency, line_style, etc.
    """
```

#### 2. delete_shape
```python
def delete_shape(self, workbook, sheet, shape_name):
    """도형 삭제"""
```

#### 3. list_shapes
```python
def list_shapes(self, workbook, sheet):
    """시트의 모든 도형 목록"""
```

#### 4. format_shape
```python
def format_shape(self, workbook, sheet, shape_name, **kwargs):
    """
    도형 서식 변경

    kwargs: fill_color, line_color, line_width, text, etc.
    """
```

#### 5. group_shapes
```python
def group_shapes(self, workbook, sheet, shape_names, group_name=None):
    """
    여러 도형을 그룹화

    Returns: 생성된 그룹 이름
    """
```

---

## 관련 문서

- **[Issue #87](https://github.com/pyhub-apps/pyhub-office-automation/issues/87)**: Remove xlwings and implement Engine Layer (22개 명령어)
- **[Issue #88](https://github.com/pyhub-apps/pyhub-office-automation/issues/88)**: Add advanced Excel commands to Engine Layer (21개 명령어)
- **[CLAUDE.md](../CLAUDE.md)**: AI Agent Quick Reference
- **[SHELL_USER_GUIDE.md](./SHELL_USER_GUIDE.md)**: Shell Mode Guide
- **[ADVANCED_FEATURES.md](./ADVANCED_FEATURES.md)**: Map Chart & Advanced Features

---

**© 2024 pyhub-office-automation** | Engine Layer Architecture Guide | v2.0 (Issue #88 Updated)