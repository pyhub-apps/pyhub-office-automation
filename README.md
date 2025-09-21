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

# 테이블 처리
oa excel table-read --output-file "data.csv"
oa excel table-write --range "A1" --data-file "data.csv"
```

### 워크북/시트 관리
```bash
oa excel workbook-create --name "새파일" --save-path "report.xlsx"
oa excel sheet-add --name "결과"
oa excel sheet-activate --name "데이터"
```

### 차트 및 피벗
```bash
oa excel chart-add --range "A1:C10" --chart-type "column"
oa excel pivot-create --source-range "A1:D100" --target-cell "F1"
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

### 3. 에러 방지 패턴
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