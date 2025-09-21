# LLM/AI 에이전트를 위한 pyhub-office-automation 사용 가이드

이 문서는 **AI 에이전트, LLM, 챗봇**이 pyhub-office-automation 도구를 효과적으로 사용하기 위한 상세 지침입니다.

## 🎯 핵심 원칙

### 1. 명령어 발견 (Command Discovery)

```bash
# 패키지 정보 및 설치 상태 확인
oa info

# Excel 명령어 전체 목록 (JSON 형식)
oa excel list --format json

# HWP 명령어 전체 목록 (JSON 형식)
oa hwp list --format json

# 특정 명령어 도움말
oa excel range-read --help
```

### 2. 컨텍스트 인식 (Context Awareness)

```bash
# 현재 열린 워크북들 확인
oa excel workbook-list --detailed

# 특정 워크북의 구조 파악
oa excel workbook-info --workbook-name "파일명.xlsx" --include-sheets
```

### 3. 에러 방지 워크플로우

**Windows (PowerShell)**
```powershell
# 1단계: 대상 워크북이 이미 열려있는지 확인
oa excel workbook-list --format json | ConvertFrom-Json | Where-Object { $_.name -eq "target.xlsx" }

# 2단계: 필요시 워크북 열기 또는 기존 연결 사용
# 열려있으면: --workbook-name 사용
# 없으면: --file-path로 열기

# 3단계: 적절한 연결 방법으로 작업 수행
```

**macOS/Linux (Bash)**
```bash
# 1단계: 대상 워크북이 이미 열려있는지 확인
oa excel workbook-list | grep "target.xlsx"

# 2단계: 필요시 워크북 열기 또는 기존 연결 사용
# 열려있으면: --workbook-name 사용
# 없으면: --file-path로 열기

# 3단계: 적절한 연결 방법으로 작업 수행
```

**AI 에이전트 권장 방법 (플랫폼 독립적)**
```bash
# 1단계: JSON 형태로 워크북 목록 조회
oa excel workbook-list --format json
# AI 에이전트가 JSON을 파싱하여 대상 워크북 존재 여부 확인

# 2단계: 존재하면 --workbook-name, 없으면 --file-path 사용
# 3단계: 적절한 연결 방법으로 작업 수행
```

## 📊 연결 방법 (Connection Methods)

모든 Excel 명령어는 3가지 워크북 연결 방법을 지원합니다:

### 방법 1: 파일 경로 (`--file-path`)

**Windows**
```bash
oa excel range-read --file-path "C:\path\to\file.xlsx" --range "A1:C10"
```

**macOS/Linux**
```bash
oa excel range-read --file-path "/path/to/file.xlsx" --range "A1:C10"
```

### 방법 2: 활성 워크북 (기본값)

```bash
oa excel range-read --range "A1:C10"
```

### 방법 3: 워크북 이름 (`--workbook-name`)

```bash
oa excel range-read --workbook-name "Sales.xlsx" --range "A1:C10"
```

## 🔄 효율적인 워크플로우 패턴

### 패턴 1: 연속 작업 (권장)

```bash
# 한번 열고 여러 작업 수행
oa excel workbook-open --file-path "report.xlsx"
oa excel sheet-add --name "Results"
oa excel range-write --range "A1" --data '["Name", "Score"]'
oa excel table-read --output-file "summary.csv"
```

### 패턴 2: 컨텍스트 기반 분석

**Windows (PowerShell)**
```powershell
# 1. 환경 파악
$workbooks = oa excel workbook-list --format json | ConvertFrom-Json

# 2. 적절한 워크북 선택 및 분석
oa excel workbook-info --workbook-name "target.xlsx" --include-sheets

# 3. 구조 기반 작업 계획 및 실행
```

**macOS/Linux (Bash)**
```bash
# 1. 환경 파악
WORKBOOKS=$(oa excel workbook-list --format json)

# 2. 적절한 워크북 선택 및 분석
oa excel workbook-info --workbook-name "target.xlsx" --include-sheets

# 3. 구조 기반 작업 계획 및 실행
```

**AI 에이전트 권장 방법 (플랫폼 독립적)**
```bash
# 1. JSON 출력을 직접 파싱
oa excel workbook-list --format json
# 결과를 AI 에이전트가 파싱하여 적절한 워크북 선택

# 2. 선택된 워크북 분석
oa excel workbook-info --workbook-name "target.xlsx" --include-sheets
```

## 🏷️ Excel Table 관리 (Windows 전용)

Excel Table(ListObject)은 **피벗테이블의 동적 범위 확장**을 가능하게 하는 핵심 기능입니다.

### 핵심 장점
- **동적 범위**: 새 데이터 추가 시 피벗테이블 범위 자동 확장
- **구조화된 참조**: 테이블명으로 범위 지정 가능
- **내장 필터**: 자동 필터/정렬 기능 제공

### Excel Table 생성 패턴

```bash
# 패턴 1: 데이터 쓰기와 동시에 Excel Table 생성 (권장)
oa excel table-write --data-file "sales.csv" --table-name "SalesData" --table-style "TableStyleMedium5"

# 패턴 2: 기존 범위를 Excel Table로 변환
oa excel table-create --range "A1:F100" --table-name "AnalysisData" --headers

# 패턴 3: Excel Table 목록 확인
oa excel table-list --detailed --format json
```

### AI 에이전트 권장 워크플로우

```bash
# 1단계: Excel Table 생성 (기존 데이터가 있는 경우)
oa excel table-create --range "A1:F100" --table-name "DataTable" --table-style "TableStyleMedium2"

# 2단계: Excel Table 확인
oa excel table-list --format json
# AI가 JSON 파싱하여 테이블 정보 확인

# 3단계: Excel Table 기반 피벗테이블 생성 (동적 범위!)
oa excel pivot-create --source-range "DataTable" --auto-position

# 💡 장점: 새 데이터가 추가되면 피벗테이블 범위가 자동으로 확장됨
```

### 플랫폼별 처리

**Windows**
```bash
# 모든 Excel Table 기능 지원
oa excel table-create --range "A1:D100" --table-name "MyTable"
```

**macOS**
```bash
# Excel Table 생성은 실패하지만 데이터 쓰기는 정상 작동
oa excel table-write --data-file "data.csv" --no-create-table  # Table 생성 비활성화
```

### 에러 처리 패턴

```bash
# AI 에이전트 권장: 플랫폼 확인 후 분기 처리
# 1. Windows에서는 Excel Table 기능 활용
# 2. macOS에서는 일반 범위로 피벗테이블 생성
```

## 📊 피벗테이블 워크플로우

피벗테이블은 **2단계 필수 과정**이 필요합니다:

### 1단계: 피벗테이블 생성 (pivot-create)

```bash
# 기본 패턴 (소스 시트명 포함)
oa excel pivot-create --source-range "Data!A1:K999" --dest-sheet "피벗분석" --dest-range "A1"

# 자동 확장 모드 (권장)
oa excel pivot-create --source-range "Data!A1" --expand "table" --dest-sheet "피벗분석" --dest-range "A1"
```

**중요 사항**:
- `--source-range`에 시트명 포함 가능 (예: "Data!A1:K999")
- 활성 시트가 아닌 경우 반드시 시트명 지정 필요
- `--expand "table"` 옵션으로 실제 데이터 영역만 자동 선택

### 2단계: 필드 설정 (pivot-configure)

```bash
# 기본 필드 설정 (간결한 형식 사용)
oa excel pivot-configure --pivot-name "PivotTable1" \
  --row-fields "카테고리,제품명" \
  --column-fields "분기" \
  --value-fields "매출액:Sum,수량:Count" \
  --filter-fields "지역"

# 간단한 설정
oa excel pivot-configure --pivot-name "PivotTable1" \
  --row-fields "제품" \
  --value-fields "매출:Sum"

# 기본값 활용 (함수 생략시 Sum 적용)
oa excel pivot-configure --pivot-name "PivotTable1" \
  --row-fields "제품" \
  --value-fields "매출"
```

### 완전한 워크플로우 예시

**Windows (PowerShell)**
```powershell
# 1. 데이터 구조 확인
oa excel range-read --range "Data!A1:K1" --format json

# 2. 피벗테이블 생성
oa excel pivot-create --source-range "Data!A1" --expand "table" --dest-sheet "피벗분석" --dest-range "A1"

# 3. 필드 설정 (간결한 형식 사용)
oa excel pivot-configure --pivot-name "PivotTable1" `
  --row-fields "카테고리,제품명" `
  --value-fields "매출액:Sum"

# 4. 데이터 새로고침
oa excel pivot-refresh --pivot-name "PivotTable1"
```

**Windows (Command Prompt)**
```cmd
# 간결한 형식 사용 (권장)
oa excel pivot-configure --pivot-name "PivotTable1" --row-fields "카테고리,제품명" --value-fields "매출액:Sum"

# JSON 형식도 지원 (복잡한 설정용)
oa excel pivot-configure --pivot-name "PivotTable1" --row-fields "카테고리,제품명" --value-fields "[{\"field\":\"매출액\",\"function\":\"Sum\"}]"
```

**macOS/Linux (Bash)**
```bash
# 1. 데이터 구조 확인
oa excel range-read --range "Data!A1:K1" --format json

# 2. 피벗테이블 생성
oa excel pivot-create --source-range "Data!A1" --expand "table" --dest-sheet "피벗분석" --dest-range "A1"

# 3. 필드 설정
oa excel pivot-configure --pivot-name "PivotTable1" \
  --row-fields "카테고리,제품명" \
  --value-fields "매출액:Sum"

# 4. 데이터 새로고침
oa excel pivot-refresh --pivot-name "PivotTable1"
```

### AI 에이전트 권장 패턴

```bash
# 1. JSON으로 헤더 정보 확인
oa excel range-read --range "Data!A1:K1" --format json
# AI가 JSON을 파싱하여 컬럼명 파악

# 2. 자동 확장으로 안전한 피벗테이블 생성
oa excel pivot-create --source-range "Data!A1" --expand "table" --dest-sheet "피벗분석" --dest-range "A1"

# 3. 파악한 컬럼명으로 필드 설정
oa excel pivot-configure --pivot-name "PivotTable1" \
  --row-fields "실제컬럼명" \
  --value-fields "숫자컬럼명:Sum"
```

### 집계 함수 옵션

```bash
# 간결한 형식에서 사용 가능한 함수들
Sum        # 합계 (기본값)
Count      # 개수
Average    # 평균
Max        # 최대값
Min        # 최소값
Product    # 곱
StdDev     # 표준편차
Var        # 분산

# 사용 예시
--value-fields "매출:Sum,수량:Count,단가:Average"
```

### 시트명 지정 규칙

- **source-range**: 시트명 포함 가능 ("시트명!범위")
- **dest-sheet**: 별도 옵션으로 지정
- **활성 시트가 아닌 경우**: 반드시 시트명 명시 필요

## 📋 데이터 처리 규칙

### JSON 입력 형식

```bash
# 단일 값
--data '"텍스트"'
--data '123'

# 1차원 배열 (행 또는 열)
--data '["값1", "값2", "값3"]'

# 2차원 배열 (테이블)
--data '[["헤더1", "헤더2"], ["데이터1", "데이터2"]]'
```

### 대용량 데이터 처리

**Windows**
```powershell
# 큰 데이터는 임시 파일 사용
oa excel table-write --data-file "C:\temp\large_data.json" --table-name "BigTable"
```

**macOS/Linux**
```bash
# 큰 데이터는 임시 파일 사용
oa excel table-write --data-file "/tmp/large_data.json" --table-name "BigTable"
```

## ⚠️ 에러 처리 가이드

### 일반적인 에러와 해결법

1. **"Workbook not found"**
   - `oa excel workbook-list`로 확인
   - 파일 경로 정확성 검증
   - 파일이 이미 다른 프로세스에서 사용 중인지 확인

2. **"Sheet not found"**
   - `oa excel workbook-info --include-sheets`로 시트 목록 확인
   - 시트 이름 정확성 검증

3. **"Range error"**
   - Excel 범위 형식 확인 (예: "A1:C10")
   - 시트 크기 내의 범위인지 확인

4. **"Permission denied"**
   - 파일이 읽기 전용인지 확인
   - 다른 프로그램에서 파일을 사용 중인지 확인

### 에러 처리 베스트 프랙티스

**Windows (PowerShell)**
```powershell
# 1. 사전 검증
$workbooks = oa excel workbook-list --format json | ConvertFrom-Json
$targetWorkbook = $workbooks.workbooks | Where-Object { $_.name -eq "target.xlsx" }

# 2. 안전한 실행
if ($targetWorkbook) {
    oa excel range-read --workbook-name "target.xlsx" --range "A1:C10"
} else {
    Write-Host "워크북이 열려있지 않습니다."
}
```

**macOS/Linux (Bash, jq 필요)**
```bash
# 1. 사전 검증
oa excel workbook-list --format json | jq '.workbooks[] | select(.name=="target.xlsx")'

# 2. 안전한 실행
if [ $? -eq 0 ]; then
    oa excel range-read --workbook-name "target.xlsx" --range "A1:C10"
else
    echo "워크북이 열려있지 않습니다."
fi
```

**AI 에이전트 권장 방법 (플랫폼 독립적)**
```bash
# 1. JSON 직접 파싱으로 검증
oa excel workbook-list --format json
# AI 에이전트가 JSON을 파싱하여 target.xlsx 존재 여부 확인

# 2. 조건부 실행
# 존재하는 경우만 명령 실행
oa excel range-read --workbook-name "target.xlsx" --range "A1:C10"
```

## 🏗️ 플랫폼 고려사항

### Windows (권장)
- **지원 기능**: Excel, HWP 모두 완전 지원
- **인터페이스**: COM 인터페이스 사용
- **성능**: 최고 성능 및 안정성
- **쉘**: PowerShell 또는 Command Prompt 사용 가능
- **경로 구분자**: 백슬래시(`\`) 또는 슬래시(`/`) 모두 지원

### macOS
- **지원 기능**: Excel만 지원 (xlwings AppleScript 모드)
- **제한사항**: HWP 미지원
- **특별 기능**: 한글 파일명 자동 정규화 지원 (NFD → NFC)
- **쉘**: Bash 또는 Zsh
- **경로 구분자**: 슬래시(`/`)

### Docker/Linux
- **지원 기능**: Excel 도구 비활성화
- **제한사항**: 제한적 기능만 사용 가능
- **쉘**: Bash
- **경로 구분자**: 슬래시(`/`)

## 📈 성능 최적화 팁

1. **연결 재사용**: 같은 워크북에 여러 작업 시 옵션 없이 활성 워크북 자동 사용 또는 `--workbook-name` 사용
2. **배치 처리**: 여러 데이터 작업을 하나의 큰 범위로 통합
3. **임시 파일**: 큰 데이터는 `--data-file` 옵션 사용
4. **컨텍스트 캐싱**: `workbook-list`, `workbook-info` 결과를 캐시하여 반복 호출 최소화
5. **플랫폼 최적화**: Windows에서 최상의 성능, macOS에서는 Excel만 사용

## 🔧 AI 에이전트 특별 권장사항

1. **플랫폼 독립적 접근**:
   - 셸 스크립트 대신 JSON 파싱을 활용
   - `--format json` 옵션을 적극 활용

2. **에러 방지**:
   - 작업 전 항상 `workbook-list`로 상황 파악
   - 파일 경로는 절대 경로 사용 권장

3. **워크플로우 최적화**:
   - 연속 작업 시 옵션 없이 활성 워크북 자동 사용
   - 필요한 경우에만 새 워크북 열기

4. **크로스 플랫폼 대응**:
   - Windows: PowerShell 문법 고려
   - macOS/Linux: Bash 문법 및 도구 가용성 확인

## 🔒 보안 및 개인정보

- **중요**: 문서 내용은 AI 학습에 사용되지 않습니다
- 임시 파일은 작업 완료 후 자동 삭제됩니다
- 로컬 처리만 수행, 외부 서버로 데이터 전송 없음
- 파일 경로 검증으로 디렉토리 순회 공격 방지

---

*이 가이드는 AI 에이전트가 안전하고 효율적으로 Office 자동화를 수행할 수 있도록 설계되었습니다.*

문의 : 파이썬사랑방 이진석 (me@pyhub.kr)