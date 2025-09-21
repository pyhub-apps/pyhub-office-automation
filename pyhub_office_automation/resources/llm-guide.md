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