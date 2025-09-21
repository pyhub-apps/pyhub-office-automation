# pyhub-office-automation

**AI 에이전트를 위한 Office 자동화 CLI 도구**

Excel과 HWP 문서를 명령줄에서 제어하는 Python 패키지입니다. JSON 출력과 구조화된 에러 처리로 AI 에이전트가 쉽게 사용할 수 있도록 설계되었습니다.

## 🤖 LLM/AI 에이전트를 위한 핵심 기능

- **구조화된 JSON 출력**: 모든 명령어가 AI 파싱에 최적화된 JSON 반환
- **스마트 연결 방법**: 옵션 없이 활성 워크북 자동 선택, `--workbook-name`으로 Excel 재실행 없이 연속 작업
- **컨텍스트 인식**: `workbook-list`로 현재 상황 파악 후 적절한 작업 수행
- **에러 방지**: 작업 전 상태 확인으로 안전한 자동화 워크플로우
- **한국 환경 최적화**: 한글 파일명 지원, HWP 자동화 (Windows)

## 🚀 빠른 시작

```bash
# 설치
pip install pyhub-office-automation

# 설치 확인
oa info

# 현재 열린 Excel 파일 확인
oa excel workbook-list --detailed

# 활성 워크북에서 데이터 읽기 (Excel이 이미 열려있는 경우)
oa excel range-read --range "A1:C10"

# 파일로 직접 접근
oa excel range-read --file-path "/path/to/file.xlsx" --range "A1:C10"
```

## 📊 핵심 Excel 명령어

### 상황 파악
```bash
oa excel workbook-list                    # 열린 파일 목록
oa excel workbook-info                     # 활성 파일 정보
oa excel workbook-info --workbook-name "파일.xlsx" --include-sheets  # 특정 파일 구조
```

### 데이터 작업
```bash
# 데이터 읽기/쓰기
oa excel range-read --range "A1:C10"
oa excel range-write --range "A1" --data '["이름", "나이", "부서"]'

# 테이블 처리 (기본)
oa excel table-read --output-file "data.csv"
oa excel table-write --range "A1" --data-file "data.csv"

# Excel Table 관리 (Windows 전용)
oa excel table-create --range "A1:D100" --table-name "SalesData"  # 범위를 Excel Table로 변환
oa excel table-list --detailed                                   # Excel Table 목록 조회
oa excel table-write --data-file "data.csv" --table-name "AutoTable"  # 데이터 쓰기 + Table 생성
```

### 워크북/시트 관리
```bash
oa excel workbook-create --name "새파일" --save-path "report.xlsx"
oa excel sheet-add --name "결과"
oa excel sheet-activate --name "데이터"
```

### 차트 생성 (두 가지 방식)

#### 정적 차트 - chart-add
```bash
# 일반 데이터 범위에서 차트 생성
# 데이터가 변경되어도 차트는 고정된 범위만 표시
oa excel chart-add --data-range "A1:C10" --chart-type "column" --title "매출 현황"

# 자동 배치로 차트 생성
oa excel chart-add --data-range "A1:C10" --auto-position --chart-type "line" --title "추세 분석"
```

#### 동적 피벗차트 - chart-pivot-create (Windows 전용)
```bash
# 기존 피벗테이블 기반으로 차트 생성
# 피벗테이블 필터/재배치 시 차트도 자동 업데이트
oa excel chart-pivot-create --pivot-name "SalesAnalysis" --chart-type "column" --title "동적 매출 분석"

# 다른 시트에 피벗차트 생성
oa excel chart-pivot-create --pivot-name "ProductSummary" --chart-type "pie" --sheet "Dashboard" --position "B2"
```

💡 **차트 선택 가이드**:
- **chart-add**: 고정 데이터, 간단한 시각화, 일회성 차트, 크로스 플랫폼
- **chart-pivot-create**: 대용량 데이터, 동적 분석, 대시보드용, Windows 전용

### 피벗테이블 생성

```bash
# 피벗테이블 생성 (2단계 필수)
# 1단계: 빈 피벗테이블 생성
# source-range에 시트명 포함 가능 (예: "Data!A1:D100")
oa excel pivot-create --source-range "Data!A1:D100" --expand "table" --dest-sheet "피벗" --dest-range "F1"

# 2단계: 필드 설정 (반드시 필요)
# 간결한 형식 사용 (권장)
oa excel pivot-configure --pivot-name "PivotTable1" \
  --row-fields "지역,제품" \
  --value-fields "매출:Sum" \
  --clear-existing
```

### 여러 객체 자동 배치 (겹침 방지)
```bash
# 첫 번째 피벗테이블 (수동 위치)
oa excel pivot-create --source-range "A1:D100" --dest-range "F1"

# 두 번째 피벗테이블 (자동 배치)
oa excel pivot-create --source-range "A1:D100" --auto-position

# 세 번째 피벗테이블 (사용자 설정)
oa excel pivot-create --source-range "A1:D100" --auto-position --spacing 3 --preferred-position "bottom"

# 정적 차트 자동 배치
oa excel chart-add --data-range "A1:C10" --auto-position --chart-type "line"

# 피벗차트 자동 배치 (Windows)
oa excel chart-pivot-create --pivot-name "PivotTable1" --chart-type "column" --sheet "Dashboard" --position "H1"

# 겹침 검사 후 생성
oa excel chart-add --data-range "A1:C10" --position "K1" --check-overlap
```

## 🔄 AI 워크플로우 예제

### 1. 스마트 상황 파악 후 작업
```bash
# 1단계: 현재 상황 파악
oa excel workbook-list

# 2단계: AI가 JSON 파싱하여 적절한 연결 방법 선택
# 파일이 열려있으면 --workbook-name, 없으면 --file-path 사용

# 3단계: 연속 작업
oa excel workbook-info --workbook-name "sales.xlsx" --include-sheets
oa excel range-read --workbook-name "sales.xlsx" --range "Sheet1!A1:Z100"
oa excel chart-add --workbook-name "sales.xlsx" --range "A1:C10"
```

### 2. 연속 데이터 처리 (리소스 효율적)
```bash
# Excel을 한 번만 열고 여러 작업 수행
oa excel workbook-open --file-path "data.xlsx"
oa excel sheet-add --name "분석결과"
oa excel range-write --sheet "분석결과" --range "A1" --data '[...]'
oa excel chart-add --sheet "분석결과" --range "A1:C10"
```

### 3. 완전한 피벗테이블 워크플로우
```bash
# 1단계: 데이터 확인
oa excel range-read --range "A1:K1"  # 헤더 확인

# 2단계: 피벗테이블 생성
oa excel pivot-create --source-range "Data!A1:K999" --expand "table" --dest-sheet "피벗분석" --dest-range "A1"

# 3단계: 필드 배치 (실제 컬럼명 사용)
oa excel pivot-configure --pivot-name "PivotTable1" \
  --row-fields "카테고리,제품명" \
  --column-fields "분기" \
  --value-fields "매출액:Sum,수량:Count" \
  --filter-fields "지역" \
  --clear-existing

# 4단계: 데이터 새로고침
oa excel pivot-refresh --pivot-name "PivotTable1"
```

### 4. Excel Table 기반 고급 피벗 워크플로우 (Windows 전용)
```bash
# 🎯 향상된 워크플로우: Excel Table → 동적 피벗테이블

# 1단계: 데이터를 Excel Table로 변환 (동적 범위 확장을 위해)
oa excel table-write --data-file "sales.csv" --table-name "SalesData" --table-style "TableStyleMedium5"

# 2단계: Excel Table 확인
oa excel table-list --detailed

# 3단계: Excel Table 기반 피벗테이블 생성 (범위 자동 확장!)
oa excel pivot-create --source-range "SalesData" --auto-position --pivot-name "SalesPivot"

# 4단계: 피벗테이블 필드 설정
oa excel pivot-configure --pivot-name "SalesPivot" \
  --row-fields "지역,제품" \
  --value-fields "매출:Sum" \
  --clear-existing

# 💡 장점: 새 데이터 추가 시 피벗테이블 범위가 자동으로 확장됨!
# 기존 범위를 Excel Table로 변환하는 경우:
oa excel table-create --range "A1:F100" --table-name "AnalysisData" --headers
```

### 5. 에러 방지 패턴
```bash
# 안전한 워크플로우: 확인 → 연결 → 작업
oa excel workbook-list | grep "target.xlsx"  # 파일 열림 확인
# 있으면: --workbook-name 사용, 없으면: --file-path로 열기
oa excel range-read --workbook-name "target.xlsx" --range "A1:C10"
```

## ✨ 특별 기능

- **자동 워크북 선택**: 옵션 없이 활성 워크북 자동 사용으로 Excel 재실행 없이 연속 작업
- **`--workbook-name`**: 파일명으로 직접 접근, 경로 불필요
- **워크북 연결 방법**: 옵션 없음(활성), `--file-path`(파일), `--workbook-name`(이름)
- **🎯 자동 배치**: 피벗테이블과 차트가 겹치지 않게 자동으로 빈 공간 찾아 배치
- **⚠️ 겹침 검사**: 지정된 위치의 충돌 여부를 사전 확인하여 경고 제공
- **JSON 최적화**: 모든 출력이 AI 에이전트 파싱에 최적화
- **한글 파일명 지원**: macOS에서 한글 자소분리 문제 자동 해결
- **37개 Excel 명령어**: 워크북/시트/데이터/차트/피벗/도형/슬라이서 전체 지원

## 📋 명령어 발견

```bash
# 전체 명령어 목록 (JSON)
oa excel list
oa hwp list

# 특정 명령어 도움말
oa excel range-read --help

# LLM 사용 가이드
oa llm-guide
```

## 🖥️ 지원 플랫폼

- **Windows 10/11**: Excel + HWP 전체 기능

---

**문의**: 파이썬사랑방 이진석 (me@pyhub.kr)