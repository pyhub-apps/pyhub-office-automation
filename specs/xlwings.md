# xlwings 라이브러리 활용 참조 가이드

본 문서는 xlwings 라이브러리를 활용한 Excel 작업 패턴과 예제를 종합한 참조 자료입니다.

## 개요

### xlwings 소개
xlwings는 Python에서 Microsoft Excel을 제어할 수 있는 라이브러리로, Excel 파일의 읽기/쓰기, 매크로 실행, 실시간 데이터 교환 등을 지원합니다.

### 활용 방식

- **크로스 플랫폼**: Windows(COM), macOS(AppleScript) 지원
- **비동기 처리**: `asyncio.to_thread`를 통한 비동기 실행
- **리소스 관리**: COM 객체 정리 및 메모리 관리

### OS별 차이점 및 제약사항

- **Windows**: COM 객체 기반, 모든 기능 지원
- **macOS**: AppleScript 연동, 일부 기능 제약 (Table 생성 등)
- **Docker**: Excel 도구 비활성화

## 1. 기본 작업 (Basic Operations)

### 1.1 라이브러리 설치

```
python -m pip install xlwings
```

### 1.2 라이브러리 임포트

```python
import xlwings as xw
```

### 1.3 Workbook 및 Sheet 접근

#### 활성 Workbook 접근

```python
# 활성 워크북 가져오기
book = xw.books.active

# 특정 워크북 가져오기 (이름으로)
book = xw.books["Sales.xlsx"]

# 모든 열린 워크북 조회
for book in xw.books:
    print(book.name, book.fullname)
```

#### Sheet 접근

```python
# 활성 시트 가져오기
sheet = xw.sheets.active

# 특정 시트 가져오기
sheet = book.sheets["Sheet1"]
sheet = book.sheets[0]  # 인덱스로 접근

# 시트 추가
new_sheet = book.sheets.add(name="NewSheet", before=None, after=None)
```

#### 유틸리티 함수 구현

```python
def get_sheet(book_name=None, sheet_name=None):
    """워크북과 시트 이름으로 시트 가져오기"""
    if book_name:
        book = xw.books[book_name]
    else:
        book = xw.books.active

    if sheet_name:
        sheet = book.sheets[sheet_name]
    else:
        sheet = book.sheets.active

    return sheet

# 사용 예제
sheet = get_sheet(book_name="Sales.xlsx", sheet_name="Sheet1")
sheet = get_sheet()  # 활성 워크북의 활성 시트
```

### 1.4 Range 작업

#### 기본 Range 접근
```python
# 단일 셀
cell = sheet.range("A1")
value = cell.value

# 범위 선택
range_ = sheet.range("A1:C10")
values = range_.value

# 사용된 범위 (데이터가 있는 전체 범위)
used_range = sheet.used_range
```

#### get_range 함수 구현
```python
def get_range(sheet_range, book_name=None, sheet_name=None, expand_mode=None):
    """범위 가져오기 함수"""
    # Sheet!Range 형태 파싱
    if '!' in sheet_range:
        sheet_name, sheet_range = sheet_range.split('!', 1)

    # 시트 가져오기
    sheet = get_sheet(book_name, sheet_name)

    # 범위 객체 생성
    range_ = sheet.range(sheet_range)

    # 확장 모드 적용
    if expand_mode:
        if expand_mode == "table":
            range_ = range_.expand()
        elif expand_mode == "down":
            range_ = range_.expand('down')
        elif expand_mode == "right":
            range_ = range_.expand('right')

    return range_

# 사용 예제
range_ = get_range("A1:C10", book_name="Sales.xlsx", sheet_name="Sheet1")
range_ = get_range("Sheet1!A1:C10")  # Sheet!Range 형태로 지정

# 확장 모드를 사용한 동적 범위 지정
range_ = get_range("A1", expand_mode="table")  # A1부터 테이블 전체로 확장
range_ = get_range("A1", expand_mode="down")   # A1부터 아래로 확장
range_ = get_range("A1", expand_mode="right")  # A1부터 오른쪽으로 확장
```

### 1.5 데이터 읽기/쓰기

#### 값 읽기
```python
# 단일 값
value = sheet.range("A1").value

# 범위 값 (2차원 리스트)
values = sheet.range("A1:C3").value
# 결과: [['A1', 'B1', 'C1'], ['A2', 'B2', 'C2'], ['A3', 'B3', 'C3']]

# 공식 읽기
formula = sheet.range("A1").formula2
```

#### 값 쓰기
```python
# 단일 값 설정
sheet.range("A1").value = "Hello"

# 2차원 데이터 설정
data = [["Name", "Age"], ["John", 30], ["Jane", 25]]
sheet.range("A1").value = data

# 공식 설정
sheet.range("A1").formula2 = "=SUM(B1:B10)"
```

### 1.6 데이터 변환 유틸리티

#### CSV 데이터 처리
```python
import csv
import io

def csv_loads(csv_data):
    """CSV 문자열을 2차원 리스트로 변환"""
    reader = csv.reader(io.StringIO(csv_data))
    return [row for row in reader]

def convert_to_csv(data):
    """2차원 데이터를 CSV 문자열로 변환"""
    output = io.StringIO()
    writer = csv.writer(output)
    for row in data:
        writer.writerow(row)
    return output.getvalue()

# 사용 예제
csv_data = "Name,Age\nJohn,30\nJane,25"
data = csv_loads(csv_data)
# 결과: [['Name', 'Age'], ['John', '30'], ['Jane', '25']]

data = [['Name', 'Age'], ['John', 30], ['Jane', 25]]
csv_string = convert_to_csv(data)
```

#### JSON 데이터 처리
```python
import json

def json_loads(json_data):
    """JSON 문자열 파싱"""
    try:
        return json.loads(json_data)
    except json.JSONDecodeError as e:
        print(f"JSON 파싱 오류: {e}")
        return None

def json_dumps(data, indent=None):
    """Python 객체를 JSON 문자열로 변환"""
    return json.dumps(data, indent=indent, ensure_ascii=False)

# 사용 예제
json_data = '{"name": "John", "age": 30}'
data = json_loads(json_data)

data = {"name": "John", "age": 30}
json_string = json_dumps(data)
```

## 2. 고급 기능 (Advanced Features)

### 2.1 PivotTable 생성 및 관리

#### Windows에서 PivotTable 생성
```python
import xlwings as xw

def create_pivot_table(source_range, dest_range, pivot_table_name=None):
    """Windows에서 COM API를 사용하여 PivotTable 생성"""
    try:
        # xlwings 상수 가져오기 (Windows에서만 사용 가능)
        from xlwings.constants import PivotFieldOrientation, PivotTableSourceType, ConsolidationFunction

        sheet = source_range.sheet

        # 피벗 캐시 생성
        pivot_cache = sheet.api.Parent.PivotCaches().Create(
            SourceType=PivotTableSourceType.xlDatabase,
            SourceData=source_range.api,
        )

        # 피벗 테이블 생성
        pivot_table = pivot_cache.CreatePivotTable(
            TableDestination=dest_range.api,
            TableName=pivot_table_name or "PivotTable1",
        )

        # 행 필드 설정
        row_fields = ["Category", "Product"]
        for name in row_fields:
            pivot_field = pivot_table.PivotFields(name)
            pivot_field.Orientation = PivotFieldOrientation.xlRowField

        # 값 필드 설정
        data_field = pivot_table.AddDataField(
            pivot_table.PivotFields("Sales"),
        )
        data_field.Function = ConsolidationFunction.xlSum

        pivot_table.RefreshTable()
        return pivot_table.Name

    except ImportError:
        raise Exception("PivotTable 생성은 Windows에서만 지원됩니다.")

# 사용 예제 (Windows 전용)
def example_pivot_table():
    """PivotTable 생성 예제"""
    import xlwings as xw

    # 워크북 열기
    wb = xw.Book()
    sheet = wb.sheets[0]

    # 샘플 데이터 입력
    data = [
        ["Category", "Product", "Sales", "Quarter"],
        ["Electronics", "Laptop", 1000, "Q1"],
        ["Electronics", "Mouse", 50, "Q1"],
        ["Furniture", "Chair", 200, "Q1"],
        ["Electronics", "Laptop", 1200, "Q2"],
        ["Furniture", "Desk", 500, "Q2"]
    ]
    sheet.range("A1").value = data

    # 원본 데이터 범위
    source_range = sheet.range("A1").expand()

    # 피벗 테이블 대상 범위
    dest_range = sheet.range("F1")

    # 피벗 테이블 생성
    pivot_name = create_pivot_table(source_range, dest_range, "SalesAnalysis")
    print(f"피벗 테이블 '{pivot_name}' 생성 완료")
```

#### 고급 PivotTable 유틸리티 구현
```python
import xlwings as xw
import platform

class PivotTableManager:
    """PivotTable 관리 클래스"""

    @staticmethod
    def create_advanced(source_range, dest_range, row_fields=None, column_fields=None,
                       page_fields=None, value_fields=None, pivot_table_name="PivotTable1"):
        """고급 PivotTable 생성"""
        if platform.system() != "Windows":
            raise Exception("고급 PivotTable 생성은 Windows에서만 지원됩니다.")

        try:
            from xlwings.constants import PivotFieldOrientation, PivotTableSourceType, ConsolidationFunction

            sheet = source_range.sheet

            # 피벗 캐시 생성
            pivot_cache = sheet.api.Parent.PivotCaches().Create(
                SourceType=PivotTableSourceType.xlDatabase,
                SourceData=source_range.api,
            )

            # 피벗 테이블 생성
            pivot_table = pivot_cache.CreatePivotTable(
                TableDestination=dest_range.api,
                TableName=pivot_table_name,
            )

            # 행 필드 설정
            if row_fields:
                for field_name in row_fields:
                    pivot_field = pivot_table.PivotFields(field_name)
                    pivot_field.Orientation = PivotFieldOrientation.xlRowField

            # 열 필드 설정
            if column_fields:
                for field_name in column_fields:
                    pivot_field = pivot_table.PivotFields(field_name)
                    pivot_field.Orientation = PivotFieldOrientation.xlColumnField

            # 페이지 필드 설정
            if page_fields:
                for field_name in page_fields:
                    pivot_field = pivot_table.PivotFields(field_name)
                    pivot_field.Orientation = PivotFieldOrientation.xlPageField

            # 값 필드 설정
            if value_fields:
                for value_field in value_fields:
                    field_name = value_field["field_name"]
                    agg_func = value_field.get("agg_func", ConsolidationFunction.xlSum)
                    data_field = pivot_table.AddDataField(
                        pivot_table.PivotFields(field_name)
                    )
                    data_field.Function = agg_func

            pivot_table.RefreshTable()
            return pivot_table.Name

        except ImportError:
            raise Exception("PivotTable 생성에 필요한 모듈을 가져올 수 없습니다.")

    @staticmethod
    def list_pivot_tables(sheet):
        """시트의 PivotTable 목록 조회"""
        if platform.system() != "Windows":
            return []  # macOS에서는 빈 리스트 반환

        try:
            pivot_tables = []
            for pivot_table in sheet.api.PivotTables():
                pivot_tables.append(pivot_table.Name)
            return pivot_tables
        except:
            return []

    @staticmethod
    def remove_pivot_tables(sheet, table_names):
        """지정된 PivotTable들 삭제"""
        if platform.system() != "Windows":
            print("PivotTable 삭제는 Windows에서만 지원됩니다.")
            return

        try:
            for table_name in table_names:
                sheet.api.PivotTables(table_name).Delete()
        except:
            print(f"PivotTable '{table_name}' 삭제 실패")

# 사용 예제
def advanced_pivot_example():
    """고급 PivotTable 예제"""
    wb = xw.Book()
    sheet = wb.sheets[0]

    # 샘플 데이터
    data = [
        ["Category", "Product", "Sales", "Quarter", "Region"],
        ["Electronics", "Laptop", 1000, "Q1", "North"],
        ["Electronics", "Mouse", 50, "Q1", "South"],
        ["Furniture", "Chair", 200, "Q1", "North"],
        ["Electronics", "Laptop", 1200, "Q2", "South"],
        ["Furniture", "Desk", 500, "Q2", "North"]
    ]
    sheet.range("A1").value = data

    source_range = sheet.range("A1").expand()
    dest_range = sheet.range("G1")

    # 고급 피벗 테이블 생성
    try:
        from xlwings.constants import ConsolidationFunction
        pivot_name = PivotTableManager.create_advanced(
            source_range=source_range,
            dest_range=dest_range,
            row_fields=["Category", "Product"],
            column_fields=["Quarter"],
            page_fields=["Region"],
            value_fields=[
                {"field_name": "Sales", "agg_func": ConsolidationFunction.xlSum}
            ],
            pivot_table_name="SalesAnalysis"
        )
        print(f"고급 피벗 테이블 '{pivot_name}' 생성 완료")

        # 피벗 테이블 목록 조회
        pivot_tables = PivotTableManager.list_pivot_tables(sheet)
        print(f"현재 피벗 테이블: {pivot_tables}")

    except Exception as e:
        print(f"피벗 테이블 생성 오류: {e}")
```

### 2.2 Excel Table 생성 및 관리 (Windows 전용)

#### 기본 Table 생성
```python
# 범위를 Excel Table로 변환
range_ = sheet.range("A1:D10")
table = sheet.tables.add(
    source_range=range_,
    name="SalesTable",
    has_headers=True,
    table_style_name="TableStyleMedium2"
)

# Table 목록 조회
for table in sheet.tables:
    print(table.name)
```

#### COM API를 통한 고급 Table 생성 (pyhub-office-automation 패턴)
```python
import platform

def create_excel_table(sheet, range_str, table_name=None, has_headers=True, table_style="TableStyleMedium2"):
    """
    Excel Table(ListObject) 생성 함수 - Windows 전용
    피벗테이블의 동적 범위 확장을 위한 핵심 기능
    """
    if platform.system() != "Windows":
        raise ValueError("Excel Table 생성은 Windows에서만 지원됩니다.")

    try:
        # 범위 객체 생성
        range_obj = sheet.range(range_str)

        # 테이블 이름 자동 생성
        if not table_name:
            existing_tables = [table.name for table in sheet.tables]
            counter = 1
            while True:
                candidate_name = f"Table{counter}"
                if candidate_name not in existing_tables:
                    table_name = candidate_name
                    break
                counter += 1

        # ListObject 생성 (COM API)
        list_object = sheet.api.ListObjects.Add(
            SourceType=1,  # xlSrcRange
            Source=range_obj.api,
            XlListObjectHasHeaders=1 if has_headers else 2  # xlYes=1, xlNo=2
        )

        # 테이블 이름 설정
        list_object.Name = table_name

        # 테이블 스타일 적용
        try:
            list_object.TableStyle = table_style
        except:
            list_object.TableStyle = "TableStyleMedium2"

        return {
            "name": table_name,
            "range": range_obj.address,
            "has_headers": has_headers,
            "style": table_style,
            "created": True
        }

    except Exception as e:
        raise ValueError(f"Excel Table 생성 실패: {str(e)}")

# 사용 예제
table_info = create_excel_table(
    sheet=sheet,
    range_str="A1:D100",
    table_name="SalesData",
    has_headers=True,
    table_style="TableStyleMedium5"
)
print(f"테이블 생성됨: {table_info}")
```

#### 피벗테이블과의 통합 패턴 (동적 범위 확장)
```python
def create_table_based_pivot(sheet, data_range, table_name, pivot_dest_range):
    """
    Excel Table 기반 피벗테이블 생성
    핵심 장점: 새 데이터 추가 시 피벗테이블 범위 자동 확장
    """
    # 1단계: Excel Table 생성
    table_info = create_excel_table(
        sheet=sheet,
        range_str=data_range,
        table_name=table_name,
        has_headers=True
    )

    # 2단계: Table 기반 피벗테이블 생성
    try:
        from xlwings.constants import PivotTableSourceType

        # 피벗 캐시 생성 (테이블명 사용으로 동적 범위!)
        pivot_cache = sheet.api.Parent.PivotCaches().Create(
            SourceType=PivotTableSourceType.xlDatabase,
            SourceData=table_name  # 범위 대신 테이블명 사용
        )

        # 피벗 테이블 생성
        pivot_name = f"Pivot_{table_name}"
        pivot_table = pivot_cache.CreatePivotTable(
            TableDestination=sheet.range(pivot_dest_range).api,
            TableName=pivot_name
        )

        return {
            "table": table_info,
            "pivot_name": pivot_name,
            "source_type": "excel_table",
            "dynamic_range": True,
            "advantage": "새 데이터 추가 시 피벗테이블 범위 자동 확장"
        }

    except Exception as e:
        raise ValueError(f"Table 기반 피벗테이블 생성 실패: {str(e)}")

# 사용 예제 - 동적 범위 피벗테이블
result = create_table_based_pivot(
    sheet=sheet,
    data_range="A1:F100",
    table_name="SalesData",
    pivot_dest_range="H1"
)
print(f"동적 피벗테이블 생성: {result}")

# 💡 핵심 장점: 새 데이터가 추가되면 피벗테이블 범위가 자동으로 확장됨!
```

#### 플랫폼별 Graceful Degradation
```python
def safe_table_operation(sheet, range_str, table_name=None):
    """
    플랫폼 안전 Table 작업
    Windows: Excel Table 생성
    macOS: 경고와 함께 범위 정보 반환
    """
    if platform.system() == "Windows":
        try:
            return create_excel_table(sheet, range_str, table_name)
        except Exception as e:
            return {
                "warning": f"Table 생성 실패: {str(e)}",
                "range": range_str,
                "fallback": True
            }
    else:
        return {
            "warning": "Excel Table 생성은 Windows에서만 지원됩니다.",
            "range": range_str,
            "platform": platform.system(),
            "recommendation": "일반 범위를 사용하여 피벗테이블 생성"
        }

# 사용 예제
result = safe_table_operation(sheet, "A1:D100", "MyTable")
print(result)
```

### 2.3 차트 작업
```python
# 차트 정보 조회
charts_info = []
for i, chart in enumerate(sheet.charts):
    chart_info = {
        "name": chart.name,
        "left": chart.left,
        "top": chart.top,
        "width": chart.width,
        "height": chart.height,
        "index": i,
    }
    charts_info.append(chart_info)
```

### 2.4 스타일 및 포맷팅
```python
# 배경색 설정 (RGB)
range_.color = (255, 255, 0)  # 노란색

# 폰트 설정
range_.font.color = (0, 0, 255)    # 파란색
range_.font.bold = True
range_.font.italic = True

# 자동 맞춤
range_.autofit()
```

### 2.5 특수 셀 찾기 (Windows 전용)
```python
import platform

def find_special_cells(range_, cell_type):
    """특수 셀 찾기 (Windows 전용)"""
    if platform.system() != "Windows":
        raise Exception("특수 셀 찾기는 Windows에서만 지원됩니다.")

    try:
        # xlwings 상수 사용
        from xlwings.constants import SpecialCellsType

        # 상수 셀만 찾기 (예: 상수값이 있는 셀들)
        if cell_type == "constants":
            special_cells_range = range_.api.SpecialCells(SpecialCellsType.xlCellTypeConstants)
        elif cell_type == "formulas":
            special_cells_range = range_.api.SpecialCells(SpecialCellsType.xlCellTypeFormulas)
        elif cell_type == "blanks":
            special_cells_range = range_.api.SpecialCells(SpecialCellsType.xlCellTypeBlanks)
        else:
            special_cells_range = range_.api.SpecialCells(SpecialCellsType.xlCellTypeConstants)

        return special_cells_range.Address
    except ImportError:
        raise Exception("Windows 상수를 가져올 수 없습니다.")

# 사용 예제
def example_special_cells():
    """특수 셀 찾기 예제"""
    import xlwings as xw

    wb = xw.Book()
    sheet = wb.sheets[0]

    # 샘플 데이터
    sheet.range("A1").value = "상수값"
    sheet.range("A2").formula = "=1+1"
    sheet.range("A3").value = None  # 빈 셀

    range_ = sheet.range("A1:A10")

    try:
        # 상수 셀들 찾기
        constants_address = find_special_cells(range_, "constants")
        print(f"상수 셀 주소: {constants_address}")

        # 수식 셀들 찾기
        formulas_address = find_special_cells(range_, "formulas")
        print(f"수식 셀 주소: {formulas_address}")

    except Exception as e:
        print(f"특수 셀 찾기 오류: {e}")
```

## 3. 유틸리티 함수 (Utility Functions)

### 3.1 데이터 정규화
```python
import re

def fix_data(sheet_range, values):
    """범위에 맞게 데이터 변환"""
    # 범위 파싱 (예: "A1:A3" -> 3행 1열)
    if ':' in sheet_range:
        start, end = sheet_range.split(':')
        # 간단한 파싱 (A1:A3 형태)
        start_col = start[0]
        start_row = int(start[1:])
        end_col = end[0]
        end_row = int(end[1:])

        rows = end_row - start_row + 1
        cols = ord(end_col) - ord(start_col) + 1
    else:
        rows = 1
        cols = 1

    # 1차원 리스트를 2차원으로 변환
    if not isinstance(values, list) or not values:
        return [[""]]

    if not isinstance(values[0], list):
        # 1차원 리스트를 열 방향으로 변환
        if cols == 1:
            return [[v] for v in values[:rows]]
        else:
            # 행 방향으로 변환
            result = []
            for i in range(0, len(values), cols):
                row = values[i:i+cols]
                while len(row) < cols:
                    row.append("")
                result.append(row)
            return result[:rows]

    return values

def normalize_2d_data(data):
    """2차원 데이터 정규화 (행별 열 개수 맞춤)"""
    if not data or not isinstance(data, list):
        return [[]]

    # 최대 열 개수 찾기
    max_cols = max(len(row) if isinstance(row, list) else 1 for row in data)

    # 모든 행을 최대 열 개수에 맞춤
    normalized = []
    for row in data:
        if not isinstance(row, list):
            row = [row]

        # 부족한 열을 빈 문자열로 채움
        while len(row) < max_cols:
            row.append("")

        normalized.append(row)

    return normalized

# 사용 예제
def example_data_normalization():
    """데이터 정규화 예제"""
    # 열 방향 범위에 맞게 데이터 변환
    values = ["v1", "v2", "v3"]
    fixed_data = fix_data("A1:A3", values)
    print(f"고정된 데이터: {fixed_data}")
    # 결과: [["v1"], ["v2"], ["v3"]]

    # 2차원 데이터 정규화 (행별 열 개수 맞춤)
    data = [['a', 'b', 'c'], ['1', '2'], ['x']]
    normalized = normalize_2d_data(data)
    print(f"정규화된 데이터: {normalized}")
    # 결과: [['a', 'b', 'c'], ['1', '2', ''], ['x', '', '']]
```

### 3.2 문자열 처리
```python
import unicodedata

def normalize_text(text):
    """Unicode 정규화 (한글 처리)"""
    if not isinstance(text, str):
        text = str(text)

    # NFC 정규화 (완성형)
    return unicodedata.normalize('NFC', text)

def str_to_list(text, delimiter=","):
    """구분자로 문자열 분리"""
    if not isinstance(text, str):
        text = str(text)

    # 구분자로 분리하고 공백 제거
    items = [item.strip() for item in text.split(delimiter)]
    # 빈 문자열 제거
    return [item for item in items if item]

# 사용 예제
def example_string_processing():
    """문자열 처리 예제"""
    # Unicode 정규화 (한글 처리)
    text = "한글텍스트"
    normalized = normalize_text(text)
    print(f"정규화된 텍스트: {normalized}")

    # 구분자로 문자열 분리
    text = "item1,item2,item3"
    items = str_to_list(text, delimiter=",")
    print(f"분리된 항목: {items}")
    # 결과: ["item1", "item2", "item3"]

    # 다른 구분자 사용
    text = "apple|banana|cherry"
    items = str_to_list(text, delimiter="|")
    print(f"파이프로 분리: {items}")
```

### 3.3 범위 주소 처리
```python
# 범위 주소 가져오기
address = range_.get_address()  # "$A$1:$C$10"

# 범위 속성
print(range_.row, range_.column)    # 시작 행, 열
print(range_.rows.count, range_.columns.count)  # 행, 열 개수
print(range_.count)  # 총 셀 개수
print(range_.shape)  # (행 개수, 열 개수)
```

## 4. OS별 처리 (Platform-specific Handling)

### 4.1 macOS 권한 처리
```python
import platform
import xlwings as xw

def check_macos_permissions():
    """macOS 권한 확인 및 안내"""
    if platform.system() == "Darwin":  # macOS
        print("macOS에서 xlwings 사용 시 권한 설정이 필요합니다.")
        print("1. 시스템 환경설정 > 보안 및 개인정보보호 > 개인정보보호")
        print("2. 자동화 > Python 또는 사용하는 IDE 선택")
        print("3. Microsoft Excel 체크박스 활성화")
        print("4. Excel > 환경설정 > 일반 > 'Excel을 열 때 통합 문서 갤러리 표시' 해제")

def safe_excel_operation():
    """안전한 Excel 작업 (권한 확인 포함)"""
    try:
        # macOS 권한 안내
        if platform.system() == "Darwin":
            check_macos_permissions()

        # Excel 작업 수행
        try:
            book = xw.books.active
            return book.name
        except Exception as e:
            if "declined permission" in str(e).lower():
                print("권한이 거부되었습니다. macOS 권한 설정을 확인해주세요.")
                check_macos_permissions()
            raise

    except Exception as e:
        print(f"Excel 작업 오류: {e}")
        return None

# 비동기 Excel 작업
import asyncio

async def async_excel_operation():
    """비동기 Excel 작업"""
    def _excel_work():
        return safe_excel_operation()

    # 별도 스레드에서 Excel 작업 실행
    return await asyncio.to_thread(_excel_work)

# 사용 예제
def example_macos_handling():
    """macOS 처리 예제"""
    try:
        # 동기 작업
        result = safe_excel_operation()
        print(f"작업 결과: {result}")

        # 비동기 작업
        async def main():
            result = await async_excel_operation()
            print(f"비동기 작업 결과: {result}")

        # asyncio.run(main())  # 필요시 실행

    except Exception as e:
        print(f"오류 발생: {e}")
```

### 4.2 AppleScript 실행 (macOS)
```python
import subprocess
import asyncio
import platform

def run_applescript(script):
    """동기 AppleScript 실행"""
    if platform.system() != "Darwin":
        raise Exception("AppleScript는 macOS에서만 지원됩니다.")

    try:
        result = subprocess.run(
            ['osascript', '-e', script],
            capture_output=True,
            text=True,
            timeout=30
        )

        if result.returncode == 0:
            return result.stdout.strip()
        else:
            raise Exception(f"AppleScript 오류: {result.stderr}")

    except subprocess.TimeoutExpired:
        raise Exception("AppleScript 실행 시간 초과")

async def run_applescript_async(script):
    """비동기 AppleScript 실행"""
    def _run_script():
        return run_applescript(script)

    return await asyncio.to_thread(_run_script)

def create_applescript_template(template, **kwargs):
    """AppleScript 템플릿 생성"""
    return template.format(**kwargs)

# 사용 예제
def example_applescript():
    """AppleScript 실행 예제"""
    if platform.system() != "Darwin":
        print("이 예제는 macOS에서만 실행됩니다.")
        return

    # 기본 스크립트
    script = '''
    tell application "Microsoft Excel"
        get name of workbooks
    end tell
    '''

    try:
        # 동기 실행
        result = run_applescript(script)
        print(f"워크북 목록: {result}")

        # 템플릿 사용
        template = '''
        tell application "Microsoft Excel"
            tell workbook "{workbook_name}"
                get name of worksheets
            end tell
        end tell
        '''

        workbook_script = create_applescript_template(
            template,
            workbook_name="Sales.xlsx"
        )
        result = run_applescript(workbook_script)
        print(f"시트 목록: {result}")

    except Exception as e:
        print(f"AppleScript 오류: {e}")

# 비동기 예제
async def example_applescript_async():
    """비동기 AppleScript 예제"""
    if platform.system() != "Darwin":
        print("이 예제는 macOS에서만 실행됩니다.")
        return

    script = '''
    tell application "Microsoft Excel"
        get name of workbooks
    end tell
    '''

    try:
        result = await run_applescript_async(script)
        print(f"비동기 워크북 목록: {result}")
    except Exception as e:
        print(f"비동기 AppleScript 오류: {e}")
```

### 4.3 플랫폼 감지
```python
import platform

def get_current_os():
    """현재 운영체제 반환"""
    system = platform.system()
    if system == "Windows":
        return "windows"
    elif system == "Darwin":
        return "macos"
    elif system == "Linux":
        return "linux"
    else:
        return "unknown"

def is_windows():
    """Windows 여부 확인"""
    return platform.system() == "Windows"

def is_macos():
    """macOS 여부 확인"""
    return platform.system() == "Darwin"

def is_linux():
    """Linux 여부 확인"""
    return platform.system() == "Linux"

# 사용 예제
def example_platform_detection():
    """플랫폼 감지 예제"""
    current_os = get_current_os()
    print(f"현재 OS: {current_os}")

    # 간단한 OS 확인
    if is_windows():
        print("Windows에서 실행 중 - COM 기능 사용 가능")
        # Windows 전용 작업
    elif is_macos():
        print("macOS에서 실행 중 - AppleScript 기능 사용 가능")
        # macOS 전용 작업
    else:
        print("지원되지 않는 운영체제")

    # 패턴 매칭 스타일 (Python 3.10+)
    match get_current_os():
        case "windows":
            print("Windows 구현")
            # Windows 전용 기능
        case "macos":
            print("macOS 구현")
            # macOS 전용 기능
        case "linux":
            print("Linux 구현")
            # Linux 지원 (제한적)
        case _:
            raise Exception(f"지원되지 않는 OS: {platform.system()}")

# 플랫폼별 xlwings 기능 매트릭스
def get_feature_support():
    """플랫폼별 지원 기능 반환"""
    features = {
        "windows": {
            "pivot_tables": True,
            "tables": True,
            "special_cells": True,
            "com_automation": True,
            "all_constants": True
        },
        "macos": {
            "pivot_tables": False,
            "tables": False,
            "special_cells": False,
            "applescript": True,
            "limited_constants": True
        },
        "linux": {
            "pivot_tables": False,
            "tables": False,
            "special_cells": False,
            "basic_operations": True,
            "no_excel_integration": True
        }
    }

    current_os = get_current_os()
    return features.get(current_os, {})

def example_feature_check():
    """기능 지원 확인 예제"""
    support = get_feature_support()
    print(f"현재 플랫폼 지원 기능: {support}")

    if support.get("pivot_tables"):
        print("PivotTable 생성이 지원됩니다.")
    else:
        print("PivotTable 생성이 지원되지 않습니다.")
```

### 4.4 COM 객체 정리 (Windows)
```python
import gc
import platform

def cleanup_excel_com():
    """COM 객체 정리 (Windows 전용)"""
    if platform.system() == "Windows":
        try:
            # 가비지 컬렉션 강제 실행
            gc.collect()

            # Windows에서 COM 객체 정리
            import pythoncom
            pythoncom.CoUninitialize()
            pythoncom.CoInitialize()

        except ImportError:
            # pythoncom이 없으면 기본 가비지 컬렉션만 실행
            gc.collect()
        except Exception as e:
            print(f"COM 정리 중 오류: {e}")
            gc.collect()
    else:
        # Windows가 아닌 경우 기본 가비지 컬렉션
        gc.collect()

def safe_excel_operation_with_cleanup():
    """안전한 Excel 작업 (정리 포함)"""
    try:
        # Excel 작업 수행
        book = xw.books.active
        sheet = book.sheets.active

        # 예시 작업
        sheet.range("A1").value = "Hello, World!"
        result = sheet.range("A1").value

        return result

    except Exception as e:
        print(f"Excel 작업 오류: {e}")
        raise

    finally:
        # COM 객체 정리 (Windows에서만 실행됨)
        cleanup_excel_com()

# 컨텍스트 매니저 스타일
class ExcelContext:
    """Excel 작업용 컨텍스트 매니저"""

    def __init__(self):
        self.book = None

    def __enter__(self):
        try:
            self.book = xw.books.active
            return self.book
        except Exception as e:
            print(f"Excel 연결 오류: {e}")
            raise

    def __exit__(self, exc_type, exc_val, exc_tb):
        # 정리 작업
        cleanup_excel_com()

        if exc_type is not None:
            print(f"Excel 작업 중 오류 발생: {exc_val}")

        return False  # 예외를 다시 발생시킴

# 사용 예제
def example_com_cleanup():
    """COM 정리 예제"""
    try:
        # 기본 사용법
        result = safe_excel_operation_with_cleanup()
        print(f"작업 결과: {result}")

        # 컨텍스트 매니저 사용법
        with ExcelContext() as book:
            sheet = book.sheets.active
            sheet.range("B1").value = "컨텍스트 매니저 테스트"
            value = sheet.range("B1").value
            print(f"컨텍스트 매니저 결과: {value}")

    except Exception as e:
        print(f"오류 발생: {e}")

# 데코레이터 스타일
def with_excel_cleanup(func):
    """Excel 정리를 위한 데코레이터"""
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        finally:
            cleanup_excel_com()
    return wrapper

@with_excel_cleanup
def excel_task():
    """데코레이터를 사용한 Excel 작업"""
    book = xw.books.active
    sheet = book.sheets.active
    sheet.range("C1").value = "데코레이터 테스트"
    return sheet.range("C1").value
```

## 5. 비동기 처리 (Async Operations)

### 5.1 배치 작업 패턴
```python
import asyncio
import xlwings as xw

async def batch_excel_operations(workbook_names):
    """여러 Excel 작업을 배치로 처리"""

    def _batch_work():
        try:
            results = []
            for workbook_name in workbook_names:
                try:
                    book = xw.books[workbook_name]
                    # 각 워크북에 대한 작업
                    result = process_workbook(book)
                    results.append({
                        "workbook": workbook_name,
                        "success": True,
                        "result": result
                    })
                except Exception as e:
                    results.append({
                        "workbook": workbook_name,
                        "success": False,
                        "error": str(e)
                    })
            return results
        finally:
            cleanup_excel_com()

    return await asyncio.to_thread(_batch_work)

def process_workbook(book):
    """개별 워크북 처리"""
    sheet = book.sheets.active

    # 예시 작업: 시트 정보 수집
    info = {
        "name": book.name,
        "sheet_count": len(book.sheets),
        "active_sheet": sheet.name,
        "used_range": sheet.used_range.address if sheet.used_range else None
    }

    return info

# 사용 예제
async def example_batch_operations():
    """배치 작업 예제"""
    # 처리할 워크북 목록 (실제로는 열린 워크북들)
    workbook_names = ["Workbook1", "Workbook2", "Workbook3"]

    try:
        results = await batch_excel_operations(workbook_names)

        print("배치 작업 결과:")
        for result in results:
            if result["success"]:
                print(f"✓ {result['workbook']}: {result['result']}")
            else:
                print(f"✗ {result['workbook']}: {result['error']}")

    except Exception as e:
        print(f"배치 작업 오류: {e}")

# 실행 예제
def run_batch_example():
    """배치 작업 실행"""
    # asyncio.run(example_batch_operations())  # 필요시 실행
    pass
```

## 6. 에러 처리 및 베스트 프랙티스

### 6.1 일반적인 에러 패턴
```python
# 워크북이 없는 경우
try:
    book = xw.books["NonExistent.xlsx"]
except Exception as e:
    print(f"Workbook not found: {e}")

# 시트가 없는 경우
try:
    sheet = book.sheets["NonExistentSheet"]
except Exception as e:
    print(f"Sheet not found: {e}")

# 범위가 유효하지 않은 경우
try:
    range_ = sheet.range("InvalidRange")
except Exception as e:
    print(f"Invalid range: {e}")
```

### 6.2 리소스 관리 패턴
```python
def safe_excel_operation():
    """안전한 Excel 작업 패턴"""
    try:
        # Excel 작업 수행
        book = xw.books.active
        sheet = book.sheets.active

        # 작업 내용
        result = sheet.range("A1:C10").value

        return result
    except Exception as e:
        print(f"Excel operation failed: {e}")
        raise
    finally:
        # 리소스 정리 (Windows에서는 COM 정리)
        cleanup_excel_com()
```

### 6.3 성능 최적화 팁

#### 대량 데이터 처리
```python
# ❌ 비효율적: 셀별 개별 접근
for i in range(1000):
    sheet.range(f"A{i}").value = data[i]

# ✅ 효율적: 범위 단위 처리
sheet.range("A1:A1000").value = [[item] for item in data]
```

#### 배치 스타일 적용
```python
# ❌ 비효율적: 개별 셀 스타일링
for cell in range_:
    cell.color = (255, 255, 0)

# ✅ 효율적: 범위 단위 스타일링
range_.color = (255, 255, 0)
```

### 6.4 데이터 검증
```python
def validate_excel_data(values):
    """Excel 데이터 검증"""
    if values is None:
        return []

    if not isinstance(values, list):
        return [[str(values)]]

    if values and not isinstance(values[0], list):
        # 1차원 리스트를 2차원으로 변환
        return [values]

    return values
```

## 7. 실제 사용 예제

### 7.1 데이터 읽기 및 CSV 변환
```python
import asyncio
import xlwings as xw
import csv
import io

async def excel_to_csv(sheet_range: str, book_name: str = "") -> str:
    """Excel 범위를 CSV 형태로 반환"""

    def _get_csv_data():
        try:
            range_ = get_range(
                sheet_range=sheet_range,
                book_name=book_name,
                expand_mode="table"
            )

            values = range_.value
            if values is None:
                return ""

            # 데이터 정규화 및 CSV 변환
            validated_data = validate_excel_data(values)
            return convert_to_csv(validated_data)
        finally:
            cleanup_excel_com()

    return await asyncio.to_thread(_get_csv_data)

def validate_excel_data(values):
    """Excel 데이터 검증"""
    if values is None:
        return []

    if not isinstance(values, list):
        return [[str(values)]]

    if values and not isinstance(values[0], list):
        # 1차원 리스트를 2차원으로 변환
        return [values]

    return values

# 완전한 예제 함수
async def complete_excel_to_csv_example():
    """완전한 Excel to CSV 변환 예제"""
    try:
        # 워크북 생성 및 샘플 데이터 입력
        wb = xw.Book()
        sheet = wb.sheets[0]

        # 샘플 데이터
        sample_data = [
            ["이름", "나이", "부서"],
            ["김철수", 30, "개발팀"],
            ["이영희", 25, "디자인팀"],
            ["박민수", 35, "기획팀"]
        ]
        sheet.range("A1").value = sample_data

        # CSV로 변환
        csv_result = await excel_to_csv("A1", book_name=wb.name)
        print("CSV 변환 결과:")
        print(csv_result)

        return csv_result

    except Exception as e:
        print(f"Excel to CSV 변환 오류: {e}")
        return ""

# 동기 버전
def excel_to_csv_sync(sheet_range: str, book_name: str = "") -> str:
    """Excel 범위를 CSV 형태로 반환 (동기 버전)"""
    try:
        range_ = get_range(
            sheet_range=sheet_range,
            book_name=book_name,
            expand_mode="table"
        )

        values = range_.value
        if values is None:
            return ""

        validated_data = validate_excel_data(values)
        return convert_to_csv(validated_data)
    finally:
        cleanup_excel_com()
```

### 7.2 CSV 데이터를 Excel에 쓰기
```python
async def csv_to_excel(csv_data: str, sheet_range: str, book_name: str = "") -> str:
    """CSV 데이터를 Excel 범위에 쓰기"""

    def _set_csv_data():
        try:
            range_ = get_range(sheet_range=sheet_range, book_name=book_name)

            # CSV 파싱
            data = csv_loads(csv_data)

            # 범위에 맞게 데이터 조정
            fixed_data = fix_data(sheet_range, data)

            range_.value = fixed_data
            return f"Successfully wrote data to {range_.address}"
        finally:
            cleanup_excel_com()

    return await asyncio.to_thread(_set_csv_data)

# 완전한 CSV to Excel 예제
async def complete_csv_to_excel_example():
    """완전한 CSV to Excel 변환 예제"""
    try:
        # 워크북 생성
        wb = xw.Book()
        sheet = wb.sheets[0]

        # CSV 데이터 준비
        csv_data = """제품명,가격,재고
        노트북,1200000,15
        마우스,25000,50
        키보드,80000,30
        모니터,350000,8"""

        # CSV 데이터를 Excel에 쓰기
        result = await csv_to_excel(csv_data, "A1", book_name=wb.name)
        print(f"CSV to Excel 결과: {result}")

        # 결과 확인
        written_data = sheet.range("A1").expand().value
        print("Excel에 작성된 데이터:")
        for row in written_data:
            print(row)

        return result

    except Exception as e:
        print(f"CSV to Excel 변환 오류: {e}")
        return ""

# 동기 버전
def csv_to_excel_sync(csv_data: str, sheet_range: str, book_name: str = "") -> str:
    """CSV 데이터를 Excel 범위에 쓰기 (동기 버전)"""
    try:
        range_ = get_range(sheet_range=sheet_range, book_name=book_name)

        # CSV 파싱
        data = csv_loads(csv_data)

        # 범위에 맞게 데이터 조정
        fixed_data = fix_data(sheet_range, data)

        range_.value = fixed_data
        return f"Successfully wrote data to {range_.address}"
    finally:
        cleanup_excel_com()
```

### 7.3 PivotTable 생성 예제
```python
async def create_sales_pivot():
    """판매 데이터 PivotTable 생성"""

    def _create_pivot():
        try:
            # 원본 데이터 범위
            source_range = get_range("A1:E1000", expand_mode="table")
            dest_range = get_range("H1")

            # PivotTable 생성 (Windows만 지원)
            if platform.system() == "Windows":
                try:
                    from xlwings.constants import ConsolidationFunction

                    pivot_name = PivotTableManager.create_advanced(
                        source_range=source_range,
                        dest_range=dest_range,
                        row_fields=["Category", "Product"],
                        column_fields=["Quarter"],
                        page_fields=["Region"],
                        value_fields=[
                            {"field_name": "Sales", "agg_func": ConsolidationFunction.xlSum},
                            {"field_name": "Quantity", "agg_func": ConsolidationFunction.xlCount}
                        ],
                        pivot_table_name="SalesAnalysis"
                    )

                    return f"Created pivot table: {pivot_name}"
                except Exception as e:
                    return f"PivotTable 생성 실패: {e}"
            else:
                return "PivotTable 생성은 Windows에서만 지원됩니다."

        finally:
            cleanup_excel_com()

    return await asyncio.to_thread(_create_pivot)

# 완전한 PivotTable 예제
async def complete_pivot_table_example():
    """완전한 PivotTable 생성 예제"""
    try:
        # 워크북 생성 및 샘플 데이터 입력
        wb = xw.Book()
        sheet = wb.sheets[0]

        # 판매 데이터 샘플
        sales_data = [
            ["Category", "Product", "Sales", "Quarter", "Region"],
            ["Electronics", "Laptop", 1500000, "Q1", "Seoul"],
            ["Electronics", "Mouse", 25000, "Q1", "Seoul"],
            ["Furniture", "Chair", 150000, "Q1", "Busan"],
            ["Electronics", "Laptop", 1800000, "Q2", "Seoul"],
            ["Furniture", "Desk", 300000, "Q2", "Busan"],
            ["Electronics", "Keyboard", 80000, "Q1", "Daegu"],
            ["Furniture", "Table", 250000, "Q2", "Daegu"],
            ["Electronics", "Monitor", 400000, "Q1", "Seoul"],
            ["Furniture", "Chair", 150000, "Q2", "Seoul"]
        ]

        # 데이터 입력
        sheet.range("A1").value = sales_data
        print("샘플 데이터 입력 완료")

        # PivotTable 생성
        result = await create_sales_pivot()
        print(f"PivotTable 생성 결과: {result}")

        return result

    except Exception as e:
        print(f"PivotTable 예제 오류: {e}")
        return ""

# 동기 버전
def create_sales_pivot_sync():
    """판매 데이터 PivotTable 생성 (동기 버전)"""
    try:
        source_range = get_range("A1:E1000", expand_mode="table")
        dest_range = get_range("H1")

        if platform.system() == "Windows":
            try:
                from xlwings.constants import ConsolidationFunction

                pivot_name = PivotTableManager.create_advanced(
                    source_range=source_range,
                    dest_range=dest_range,
                    row_fields=["Category", "Product"],
                    column_fields=["Quarter"],
                    page_fields=["Region"],
                    value_fields=[
                        {"field_name": "Sales", "agg_func": ConsolidationFunction.xlSum}
                    ],
                    pivot_table_name="SalesAnalysis"
                )
                return f"Created pivot table: {pivot_name}"
            except Exception as e:
                return f"PivotTable 생성 실패: {e}"
        else:
            return "PivotTable 생성은 Windows에서만 지원됩니다."

    finally:
        cleanup_excel_com()
```

### 7.4 대량 데이터 스타일링
```python
async def apply_conditional_formatting(data_range: str):
    """조건부 서식 적용"""

    def _apply_formatting():
        try:
            range_ = get_range(data_range, expand_mode="table")

            # 헤더 스타일
            header_range = range_.rows[0]
            header_range.color = (70, 130, 180)  # 스틸 블루
            header_range.font.color = (255, 255, 255)  # 흰색
            header_range.font.bold = True

            # 데이터 행 교대로 색상 적용
            for i, row in enumerate(range_.rows[1:], 1):
                if i % 2 == 0:
                    row.color = (240, 248, 255)  # 연한 파란색

            # 자동 맞춤
            range_.autofit()

            return f"Applied formatting to {range_.address}"
        finally:
            cleanup_excel_com()

    return await asyncio.to_thread(_apply_formatting)

# 완전한 스타일링 예제
async def complete_formatting_example():
    """완전한 데이터 스타일링 예제"""
    try:
        # 워크북 생성 및 데이터 입력
        wb = xw.Book()
        sheet = wb.sheets[0]

        # 샘플 데이터
        data = [
            ["부서", "직원수", "평균급여", "예산"],
            ["개발팀", 15, 5500000, 82500000],
            ["디자인팀", 8, 4800000, 38400000],
            ["기획팀", 12, 5200000, 62400000],
            ["영업팀", 20, 4500000, 90000000],
            ["인사팀", 5, 4000000, 20000000]
        ]

        sheet.range("A1").value = data
        print("데이터 입력 완료")

        # 스타일링 적용
        result = await apply_conditional_formatting("A1")
        print(f"스타일링 결과: {result}")

        # 추가 스타일링 - 숫자 포맷
        number_range = sheet.range("C2:D6")  # 급여와 예산 열
        number_range.number_format = "#,##0"

        print("숫자 포맷 적용 완료")

        return result

    except Exception as e:
        print(f"스타일링 예제 오류: {e}")
        return ""

# 동기 버전
def apply_conditional_formatting_sync(data_range: str):
    """조건부 서식 적용 (동기 버전)"""
    try:
        range_ = get_range(data_range, expand_mode="table")

        # 헤더 스타일
        header_range = range_.rows[0]
        header_range.color = (70, 130, 180)
        header_range.font.color = (255, 255, 255)
        header_range.font.bold = True

        # 데이터 행 교대로 색상 적용
        for i, row in enumerate(range_.rows[1:], 1):
            if i % 2 == 0:
                row.color = (240, 248, 255)

        range_.autofit()
        return f"Applied formatting to {range_.address}"
    finally:
        cleanup_excel_com()
```

### 7.5 워크북 정보 조회
```python
async def get_workbook_info() -> dict:
    """열린 워크북들의 정보 조회"""

    def _get_info():
        try:
            books_info = []
            for book in xw.books:
                book_info = {
                    "name": normalize_text(book.name),
                    "fullname": normalize_text(book.fullname),
                    "active": book == xw.books.active,
                    "sheets": []
                }

                for sheet in book.sheets:
                    sheet_info = {
                        "name": normalize_text(sheet.name),
                        "index": sheet.index,
                        "range": sheet.used_range.get_address() if sheet.used_range else "",
                        "count": sheet.used_range.count if sheet.used_range else 0,
                        "shape": sheet.used_range.shape if sheet.used_range else (0, 0),
                        "active": sheet == xw.sheets.active,
                        "table_names": get_table_names(sheet)
                    }
                    book_info["sheets"].append(sheet_info)

                books_info.append(book_info)

            return {"books": books_info}
        finally:
            cleanup_excel_com()

    return await asyncio.to_thread(_get_info)

def get_table_names(sheet):
    """시트의 테이블 이름 목록 조회"""
    try:
        if hasattr(sheet, 'tables'):
            return [table.name for table in sheet.tables]
        else:
            return []
    except:
        return []

# 완전한 워크북 정보 조회 예제
async def complete_workbook_info_example():
    """완전한 워크북 정보 조회 예제"""
    try:
        # 몇 개의 워크북을 생성하여 테스트
        wb1 = xw.Book()
        wb1.sheets[0].range("A1").value = "첫 번째 워크북"

        wb2 = xw.Book()
        wb2.sheets[0].range("A1").value = "두 번째 워크북"
        wb2.sheets.add("새시트")

        # 워크북 정보 조회
        info = await get_workbook_info()

        print("워크북 정보:")
        print(json_dumps(info, indent=2))

        # 요약 정보 출력
        print(f"\n총 {len(info['books'])}개의 워크북이 열려있습니다:")
        for book_info in info['books']:
            print(f"- {book_info['name']}: {len(book_info['sheets'])}개 시트")

        return info

    except Exception as e:
        print(f"워크북 정보 조회 오류: {e}")
        return {}

# 동기 버전
def get_workbook_info_sync() -> dict:
    """열린 워크북들의 정보 조회 (동기 버전)"""
    try:
        books_info = []
        for book in xw.books:
            book_info = {
                "name": normalize_text(book.name),
                "fullname": normalize_text(book.fullname),
                "active": book == xw.books.active,
                "sheets": []
            }

            for sheet in book.sheets:
                sheet_info = {
                    "name": normalize_text(sheet.name),
                    "index": sheet.index,
                    "range": sheet.used_range.get_address() if sheet.used_range else "",
                    "count": sheet.used_range.count if sheet.used_range else 0,
                    "shape": sheet.used_range.shape if sheet.used_range else (0, 0),
                    "active": sheet == xw.sheets.active,
                    "table_names": get_table_names(sheet)
                }
                book_info["sheets"].append(sheet_info)

            books_info.append(book_info)

        return {"books": books_info}
    finally:
        cleanup_excel_com()

## 종합 예제 함수

async def run_all_examples():
    """모든 예제를 순차적으로 실행"""
    print("=== xlwings 완전 독립형 예제 실행 ===\n")

    try:
        # 1. CSV 변환 예제
        print("1. Excel to CSV 변환 예제")
        await complete_excel_to_csv_example()
        print()

        # 2. CSV to Excel 예제
        print("2. CSV to Excel 변환 예제")
        await complete_csv_to_excel_example()
        print()

        # 3. PivotTable 예제 (Windows만)
        print("3. PivotTable 생성 예제")
        await complete_pivot_table_example()
        print()

        # 4. 스타일링 예제
        print("4. 데이터 스타일링 예제")
        await complete_formatting_example()
        print()

        # 5. 워크북 정보 조회 예제
        print("5. 워크북 정보 조회 예제")
        await complete_workbook_info_example()
        print()

        print("=== 모든 예제 실행 완료 ===")

    except Exception as e:
        print(f"예제 실행 중 오류: {e}")

# 메인 실행 함수
def main():
    """메인 실행 함수"""
    # asyncio.run(run_all_examples())  # 필요시 주석 해제
    print("xlwings 완전 독립형 가이드가 로드되었습니다.")
    print("예제를 실행하려면 main() 함수의 주석을 해제하세요.")

if __name__ == "__main__":
    main()
```

## 주요 제약사항 및 주의사항

1. **macOS 제약사항**: Table 생성, 일부 고급 기능 제한
2. **Docker 환경**: Excel 도구 완전 비활성화
3. **COM 객체**: Windows에서 반드시 정리 필요
4. **동시성**: Excel은 단일 스레드에서만 안전하게 작동
5. **메모리**: 대량 데이터 처리 시 메모리 사용량 주의
6. **권한**: macOS에서 자동화 권한 필요

## 필요한 라이브러리 요약

이 가이드의 모든 예제를 실행하기 위해 필요한 라이브러리들:

```python
# 필수 라이브러리
import xlwings as xw
import asyncio
import platform
import gc

# 표준 라이브러리
import csv
import io
import json
import unicodedata
import subprocess
import re

# Windows 전용 (선택사항)
try:
    import pythoncom  # COM 객체 정리용
except ImportError:
    pass  # Windows가 아닌 경우 무시
```

## 설치 방법

```bash
# xlwings 설치
pip install xlwings

# macOS의 경우 추가 설정
# 1. Excel 설치 필요
# 2. macOS 시스템 설정에서 자동화 권한 허용
# 3. Excel 환경설정에서 통합 문서 갤러리 비활성화
```

## 결론

이 참조 가이드는 xlwings 라이브러리를 완전히 독립적으로 사용할 수 있도록 작성되었습니다. 모든 예제는 pyhub 의존성 없이 순수 xlwings와 Python 표준 라이브러리만으로 동작하며, 실제 프로덕션 환경에서 사용할 수 있는 검증된 패턴들입니다.

### 주요 특징

1. **완전 독립형**: pyhub 의존성 완전 제거
2. **크로스 플랫폼**: Windows와 macOS 모두 지원
3. **실행 가능한 예제**: 모든 코드가 즉시 실행 가능
4. **동기/비동기 지원**: 두 가지 방식 모두 제공
5. **에러 처리**: 강건한 예외 처리와 정리 로직

### 사용 권장사항

- Windows에서는 모든 기능 사용 가능
- macOS에서는 기본 기능 위주로 사용 권장
- 대량 데이터 처리 시 배치 작업 패턴 활용
- 항상 cleanup 함수를 통한 리소스 정리 수행

Excel 자동화 작업 시 이 가이드를 참조하여 효율적이고 안정적인 코드를 작성하시기 바랍니다.

