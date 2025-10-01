# Excel Shell Mode - 기본 워크플로우 예제

이 문서는 Excel Shell Mode의 실제 사용 예제를 단계별로 설명합니다.

## 예제 1: 데이터 탐색 및 분석

### 시나리오
Sales.xlsx 파일에서 데이터를 탐색하고 간단한 차트를 생성합니다.

### 실행 방법

```bash
# Shell 시작
$ oa excel shell

# 1단계: 환경 파악
[Excel: None > None] > workbook-list
# 출력: 현재 열린 모든 Excel 파일 목록

[Excel: None > None] > use workbook "Sales.xlsx"
✓ Workbook set: Sales.xlsx
✓ Active sheet: Sheet1

# 2단계: 데이터 구조 확인
[Excel: Sales.xlsx > Sheet1] > sheets
Available sheets in Sales.xlsx:
  1. Sheet1 (Active)
  2. RawData
  3. Summary

[Excel: Sales.xlsx > Sheet1] > use sheet RawData
✓ Active sheet: RawData

# 3단계: 테이블 구조 파악
[Excel: Sales.xlsx > RawData] > table-list
# 출력: 테이블 이름, 컬럼, 샘플 데이터

# 4단계: 샘플 데이터 확인
[Excel: Sales.xlsx > RawData] > range-read --range A1:F5
# 출력: 처음 5행의 데이터

# 5단계: 차트 생성 시트로 이동
[Excel: Sales.xlsx > RawData] > use sheet Summary
✓ Active sheet: Summary

# 6단계: 차트 생성
[Excel: Sales.xlsx > Summary] > chart-add --data-range "RawData!A1:B10" --chart-type "Column" --title "월별 매출"
✓ Chart created: Chart1

# 7단계: 완료 및 종료
[Excel: Sales.xlsx > Summary] > exit
Goodbye!
```

### 학습 포인트
- ✅ `workbook-list`로 환경 먼저 파악
- ✅ `use` 명령으로 컨텍스트 설정
- ✅ `sheets`로 구조 확인
- ✅ `table-list`로 데이터 개요 파악
- ✅ 컨텍스트 자동 주입으로 명령어 단축

---

## 예제 2: 다중 시트 데이터 처리

### 시나리오
분기별 시트(Q1, Q2, Q3, Q4)에서 데이터를 추출하여 통합 요약 시트를 생성합니다.

### 실행 방법

```bash
$ oa excel shell --workbook-name "Quarterly_Report.xlsx"

# 1단계: 모든 시트 확인
[Excel: Quarterly_Report.xlsx > Sheet1] > sheets
Available sheets:
  1. Q1
  2. Q2
  3. Q3
  4. Q4
  5. Summary

# 2단계: Q1 데이터 확인
[Excel: Quarterly_Report.xlsx > Sheet1] > use sheet Q1
[Excel: Quarterly_Report.xlsx > Q1] > range-read --range A1:D50
# 데이터 구조 파악

# 3단계: Q2 데이터 확인
[Excel: Quarterly_Report.xlsx > Q1] > use sheet Q2
[Excel: Quarterly_Report.xlsx > Q2] > range-read --range A1:D50

# 4단계: Q3 데이터 확인
[Excel: Quarterly_Report.xlsx > Q2] > use sheet Q3
[Excel: Quarterly_Report.xlsx > Q3] > range-read --range A1:D50

# 5단계: Q4 데이터 확인
[Excel: Quarterly_Report.xlsx > Q3] > use sheet Q4
[Excel: Quarterly_Report.xlsx > Q4] > range-read --range A1:D50

# 6단계: Summary 시트로 이동하여 통합 데이터 작성
[Excel: Quarterly_Report.xlsx > Q4] > use sheet Summary
[Excel: Quarterly_Report.xlsx > Summary] > range-write --range A1 --data '[["Quarter","Revenue"],["Q1",1000],["Q2",1200],["Q3",1500],["Q4",1800]]'

# 7단계: 차트 생성
[Excel: Quarterly_Report.xlsx > Summary] > chart-add --data-range "A1:B5" --chart-type "Line" --title "분기별 매출 추이"

[Excel: Quarterly_Report.xlsx > Summary] > exit
```

### 학습 포인트
- ✅ 시트 전환으로 동일 패턴 작업 반복
- ✅ `use sheet` 명령만으로 빠른 전환
- ✅ 컨텍스트 유지로 --workbook-name 불필요
- ✅ 데이터 비교 및 통합 작업 효율적

---

## 예제 3: 피벗테이블 생성 (Windows)

### 시나리오
원본 데이터에서 피벗테이블과 피벗차트를 생성합니다.

### 실행 방법

```bash
$ oa excel shell

[Excel: None > None] > use workbook "SalesData.xlsx"
[Excel: SalesData.xlsx > None] > use sheet RawData

# 1단계: 데이터 구조 확인
[Excel: SalesData.xlsx > RawData] > range-read --range A1:A1
# 헤더 확인: Region, Product, Category, Sales, Quantity

# 2단계: 피벗 시트 생성
[Excel: SalesData.xlsx > RawData] > sheet-add --name "Pivot Analysis"
[Excel: SalesData.xlsx > RawData] > use sheet "Pivot Analysis"

# 3단계: 피벗테이블 생성
[Excel: SalesData.xlsx > Pivot Analysis] > pivot-create --source-range "RawData!A1:E1000" --expand table --dest-range "A1"
✓ Pivot table created: PivotTable1

# 4단계: 피벗테이블 필드 설정
[Excel: SalesData.xlsx > Pivot Analysis] > pivot-configure --pivot-name "PivotTable1" --row-fields "Region,Product" --value-fields "Sales:Sum" --clear-existing
✓ Pivot fields configured

# 5단계: 데이터 새로고침
[Excel: SalesData.xlsx > Pivot Analysis] > pivot-refresh --pivot-name "PivotTable1"
✓ Pivot table refreshed

# 6단계: 피벗차트 생성
[Excel: SalesData.xlsx > Pivot Analysis] > chart-pivot-create --pivot-name "PivotTable1" --chart-type "Column" --title "지역별 매출"

[Excel: SalesData.xlsx > Pivot Analysis] > exit
```

### 학습 포인트
- ✅ 복잡한 피벗 작업도 Shell에서 연속 실행
- ✅ 시트 전환으로 구조화된 분석
- ✅ 피벗테이블 → 설정 → 차트 순차 진행
- ✅ 에러 발생 시 즉시 재시도 가능

---

## 예제 4: Tab 자동완성 활용

### 시나리오
명령어를 정확히 기억하지 못할 때 Tab 자동완성을 활용합니다.

### 실행 방법

```bash
$ oa excel shell

# Tab 키로 명령어 탐색
[Excel: None > None] > wo<TAB>
# 자동완성: workbook-list, workbook-info, workbook-open, workbook-create

[Excel: None > None] > workbook-<TAB>
# 하위 명령 확인

[Excel: None > None] > workbook-list

[Excel: None > None] > use <TAB>
# 자동완성: use workbook, use sheet

[Excel: None > None] > use w<TAB>
# 자동완성: use workbook

[Excel: None > None] > use workbook "test.xlsx"

[Excel: test.xlsx > None] > sh<TAB>
# 자동완성: sheets, sheet-add, sheet-activate, sheet-delete, sheet-rename, show

[Excel: test.xlsx > None] > sheets

[Excel: test.xlsx > None] > ra<TAB>
# 자동완성: range-read, range-write, range-convert

[Excel: test.xlsx > None] > range-read --range A1:C10
```

### 학습 포인트
- ✅ Tab 키로 52개 명령어 모두 탐색 가능
- ✅ 부분 입력 후 Tab으로 자동완성
- ✅ 명령어 오타 방지
- ✅ 명령어를 정확히 몰라도 탐색 가능

---

## 예제 5: 에러 복구 패턴

### 시나리오
잘못된 명령 입력 후 빠르게 수정하여 재시도합니다.

### 실행 방법

```bash
$ oa excel shell

[Excel: None > None] > use workbook "Report.xlsx"
[Excel: Report.xlsx > None] > use sheet "Data"
[Excel: Report.xlsx > Data] > range-read --range "A1:Z100"

# 에러 발생!
Error: Sheet 'Data' not found

# 즉시 시트 목록 확인
[Excel: Report.xlsx > Data] > sheets
Available sheets:
  1. Sheet1
  2. RawData
  3. Summary

# 올바른 시트명으로 전환
[Excel: Report.xlsx > Data] > use sheet "RawData"
✓ Active sheet: RawData

# 명령 재시도 (위 화살표로 이전 명령 불러오기)
[Excel: Report.xlsx > RawData] > <UP ARROW>
[Excel: Report.xlsx > RawData] > range-read --range "A1:Z100"
# 성공!

[Excel: Report.xlsx > RawData] > exit
```

### 학습 포인트
- ✅ 세션 유지로 빠른 에러 복구
- ✅ `sheets` 명령으로 올바른 이름 확인
- ✅ 위/아래 화살표로 명령 히스토리 활용
- ✅ 컨텍스트 수정만으로 재시도 가능

---

## Shell Mode vs 일반 CLI 비교

### 동일한 작업을 두 방식으로 비교

**Shell Mode (권장)**:
```bash
$ oa excel shell
[Excel: None > None] > use workbook "sales.xlsx"
[Excel: sales.xlsx > None] > use sheet "Data"
[Excel: sales.xlsx > Data] > table-list
[Excel: sales.xlsx > Data] > range-read --range A1:C10
[Excel: sales.xlsx > Data] > chart-add --data-range "A1:B10" --chart-type "Column"
[Excel: sales.xlsx > Data] > exit
```
**입력 문자 수**: ~200자

**일반 CLI Mode**:
```bash
$ oa excel workbook-list
$ oa excel workbook-info --workbook-name "sales.xlsx"
$ oa excel table-list --workbook-name "sales.xlsx" --sheet "Data"
$ oa excel range-read --workbook-name "sales.xlsx" --sheet "Data" --range A1:C10
$ oa excel chart-add --workbook-name "sales.xlsx" --sheet "Data" --data-range "A1:B10" --chart-type "Column"
```
**입력 문자 수**: ~350자

**결과**: Shell Mode가 43% 더 짧음! ✅

---

## 추가 팁

### 1. show context 활용
현재 상태를 주기적으로 확인하여 실수 방지:
```bash
[Excel: sales.xlsx > Data] > show context
Current Context:
  Workbook: sales.xlsx
  Sheet: Data
  All Excel commands will use this context automatically.
```

### 2. clear 명령으로 화면 정리
긴 출력 후 화면 정리:
```bash
[Excel: sales.xlsx > Data] > clear
```

### 3. help로 명령어 카테고리 확인
```bash
[Excel: sales.xlsx > Data] > help

Shell Commands (8):
  - help, show, use, clear, exit, quit, sheets, workbook-info

Excel Commands by Category:
  Range (3): range-read, range-write, range-convert
  Workbook (5): workbook-list, workbook-open, ...
  Sheet (4): sheet-activate, sheet-add, ...
  Table (5): table-read, table-write, ...
  Chart (7): chart-add, chart-configure, ...
  ... (continues)
```

### 4. 명령어 히스토리 검색
- **위/아래 화살표**: 이전 명령 탐색
- **Ctrl+R**: 히스토리 검색 (reverse-i-search)

---

## 문제 해결

### Q: Shell이 시작되지 않아요
**A**: Excel이 설치되어 있고 실행 가능한지 확인하세요:
```bash
$ oa excel workbook-list
```

### Q: 컨텍스트가 자동 주입되지 않아요
**A**: `show context`로 현재 상태를 확인하세요. 워크북/시트가 설정되어 있지 않으면 `use` 명령으로 설정하세요.

### Q: Tab 자동완성이 작동하지 않아요
**A**: prompt-toolkit이 제대로 설치되었는지 확인하세요:
```bash
$ pip list | grep prompt-toolkit
```

### Q: 명령어를 잘못 입력했어요
**A**: 위 화살표로 이전 명령을 불러와서 수정하거나, 그냥 다시 입력하세요. Shell 세션은 유지됩니다.

---

## 다음 단계

- [ ] 더 복잡한 피벗테이블 시나리오 연습
- [ ] 여러 워크북 간 데이터 이동
- [ ] 매크로 실행 및 VBA 코드 실행
- [ ] 대용량 데이터 처리 (페이징, 샘플링)

**Shell Mode를 마스터하여 Excel 자동화 생산성을 10배 향상시키세요!** 🚀
