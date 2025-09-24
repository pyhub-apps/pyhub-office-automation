# pyhub-office-automation

**AI 에이전트를 위한 Office 자동화 CLI 도구**

Excel과 HWP 문서를 명령줄에서 제어하는 Python 패키지입니다. JSON 출력과 구조화된 에러 처리로 AI 에이전트가 쉽게 사용할 수 있도록 설계되었습니다.

## 🤖 LLM/AI 에이전트를 위한 핵심 기능

- **구조화된 JSON 출력**: 모든 명령어가 AI 파싱에 최적화된 JSON 반환
- **스마트 연결 방법**: 옵션 없이 활성 워크북 자동 선택, `--workbook-name`으로 Excel 재실행 없이 연속 작업
- **컨텍스트 인식**: `workbook-list`로 현재 상황 파악 후 적절한 작업 수행
- **🆕 테이블 구조 즉시 파악**: `table-list`로 컬럼명+샘플데이터를 한 번에 제공, 추가 API 호출 불필요
- **에러 방지**: 작업 전 상태 확인으로 안전한 자동화 워크플로우
- **한국 환경 최적화**: 한글 파일명 지원, HWP 자동화 (Windows)


## 🚀 빠른 시작

```bash
# 설치
pip install pyhub-office-automation

# 설치 확인
oa info

# 현재 열린 Excel 파일 확인 (상세 정보 포함)
oa excel workbook-list

# 🆕 테이블 구조와 샘플 데이터 즉시 파악 (AI 에이전트 최적화)
oa excel table-list

# 🔥 발견한 테이블 데이터 읽기 (완전한 table-driven 워크플로우)
oa excel table-read --table-name "GameData" --limit 100

# 대용량 데이터 처리 (페이징과 샘플링)
oa excel table-read --table-name "GameData" --offset 500 --limit 100
oa excel table-read --table-name "GameData" --limit 50 --sample-mode

# 특정 컬럼만 선택하여 Context 절약
oa excel table-read --table-name "GameData" --columns "게임명,글로벌 판매량" --limit 100

# 기존 방식: 일반 셀 범위 읽기 (Excel Table 외부 데이터용)
oa excel range-read --range "A1:C10"
```

## 🤖 AI 코드 어시스턴트 설정

각 AI 에이전트에 최적화된 설정 파일을 자동으로 생성합니다:

### 지원 대상
- **Claude Code** → CLAUDE.md 업데이트/생성
- **Gemini CLI** → GEMINI.md 생성
- **Codex CLI** → AGENTS.md 생성
- **모든 AI** → 전체 파일 생성

### 사용법
```bash
# Claude Code 사용자
oa ai-setup claude

# Gemini CLI 사용자
oa ai-setup gemini

# Codex CLI 사용자
oa ai-setup codex

# 모든 AI 지원 파일 생성
oa ai-setup all
```

### 자동 생성되는 내용
설정 파일에는 다음 내용이 포함됩니다:

- ✅ **`oa` 명령어 사용 가이드**: 기본 사용법과 권장 패턴
- ✅ **Python 경로 자동 탐지**: 설치된 Python 환경 자동 설정
- ✅ **Excel 데이터 차트 제안 템플릿**: 데이터 유형별 차트 추천
- ✅ **에러 처리 및 디버깅 가이드**: 자주 발생하는 문제 해결법
- ✅ **워크북 연결 최적화**: 효율적인 Excel 파일 접근 방법

### Python 환경 자동 탐지
시스템의 Python 설치를 자동으로 감지하여 지침에 포함:
- PATH 환경변수 확인
- 일반적인 설치 경로 스캔 (`anaconda3`, `Programs/Python` 등)
- matplotlib 한글 폰트 설정 (Malgun Gothic, 300dpi)

### 예시 출력
```
✅ AI 에이전트 설정 완료!
- 파일 생성: GEMINI.md
- Python 경로 감지: C:\Users\user\anaconda3\python.exe
- 차트 템플릿 추가: 5개 예시
- 다음 명령으로 확인: cat GEMINI.md

💡 사용 팁: AI 에이전트에서 이 파일을 자동으로 읽도록 설정하세요.
```

> **참고**: 이 기능은 [GitHub Issue #56](https://github.com/pyhub-apps/pyhub-office-automation/issues/56)으로 계획되어 있으며, 향후 업데이트에서 구현될 예정입니다.

## 📊 핵심 Excel 명령어

### 상황 파악
```bash
oa excel workbook-list                    # 열린 파일 목록 (상세 정보 포함)
oa excel workbook-info                     # 활성 파일 정보 (모든 정보 포함)
oa excel workbook-info --workbook-name "파일.xlsx"  # 특정 파일 구조 (모든 정보 포함)
```

### 데이터 작업
```bash
# 데이터 읽기/쓰기/변환
oa excel range-read --range "A1:C10"
oa excel range-write --range "A1" --data '["이름", "나이", "부서"]'
oa excel range-convert --range "A1:C10"  # 문자열→숫자 자동 변환

# 형식 변환 상세 옵션
oa excel range-convert --range "A1:Z100" --remove-comma  # "1,234" → 1234
oa excel range-convert --range "B2:B100" --remove-currency  # "₩1,000" → 1000
oa excel range-convert --range "C1:C50" --parse-percent  # "50%" → 0.5
oa excel range-convert --range "D1:D100" --expand table --no-save  # 테이블 전체 변환, 저장 안 함

# 테이블 처리
oa excel table-list                           # 🆕 모든 테이블 구조+샘플 데이터 (AI 최적화)
oa excel table-read --output-file "data.csv"
oa excel table-write --range "A1" --data-file "data.csv"
oa excel table-analyze --table-name "Sales"  # 🆕 특정 테이블 메타데이터 생성
oa excel metadata-generate                    # 🆕 모든 테이블 메타데이터 일괄 생성

# Excel Table 관리 (Windows 전용)
oa excel table-create --range "A1:D100" --table-name "SalesData"  # 범위를 Excel Table로 변환
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

### 데이터 분석 및 변환 (피벗 준비)

```bash
# 데이터 구조 분석 (피벗테이블 준비상태 확인)
oa excel data-analyze --range "A1:Z100" --expand "table"
oa excel data-analyze --file-path "report.xlsx" --range "Sheet1!A1:K1000"

# 데이터 변환 (피벗테이블용 형식으로 변환)
# 교차표를 세로 형식으로 변환
oa excel data-transform --source-range "A1:M100" --transform-type "unpivot" --output-sheet "PivotReady"

# 병합된 셀 해제 및 값 채우기
oa excel data-transform --source-range "A1" --expand "table" --transform-type "unmerge"

# 모든 필요한 변환 자동 적용
oa excel data-transform --source-range "Data!A1:K999" --transform-type "auto" --expand "table"

# 다단계 헤더를 단일 헤더로 결합
oa excel data-transform --source-range "A1:J50" --transform-type "flatten-headers" --output-sheet "CleanData"
```

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
oa excel workbook-info --workbook-name "sales.xlsx"  # 모든 정보 자동 포함
oa excel range-read --workbook-name "sales.xlsx" --range "Sheet1!A1:Z100"  # 값과 공식 자동 포함
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

### 3. 스마트 데이터 준비 및 피벗테이블 생성 (AI 지원)
```bash
# 1단계: 데이터 구조 자동 분석 (AI 데이터 패턴 감지)
oa excel data-analyze --range "A1:Z1000" --expand "table"
# → AI가 교차표, 병합셀, 다단계헤더, 소계혼재, 넓은형식 등 5가지 문제 자동 감지
# → 피벗테이블 준비도 평가 및 권장 변환방법 제시

# 2단계: AI 권장사항에 따른 자동 데이터 변환
oa excel data-transform --source-range "A1:Z1000" --transform-type "auto" --expand "table" --output-sheet "PivotReady"
# → AI가 감지한 모든 문제를 올바른 순서로 자동 해결
# → 병합셀 해제 → 소계 제거 → 헤더 정리 → 교차표 변환 순으로 적용

# 3단계: 변환된 데이터로 피벗테이블 생성
oa excel pivot-create --source-range "PivotReady!A1:F5000" --expand "table" --dest-sheet "분석결과" --dest-range "A1"

# 4단계: 필드 배치 (변환된 헤더 사용)
oa excel pivot-configure --pivot-name "PivotTable1" \
  --row-fields "카테고리,제품명" \
  --column-fields "측정항목" \
  --value-fields "값:Sum" \
  --clear-existing

# 5단계: 데이터 새로고침
oa excel pivot-refresh --pivot-name "PivotTable1"
```

### 4. 완전한 피벗테이블 워크플로우 (기존 방식)
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

### 5. Excel Table 기반 고급 피벗 워크플로우 (Windows 전용)
```bash
# 🎯 향상된 워크플로우: Excel Table → 동적 피벗테이블

# 1단계: 데이터를 Excel Table로 변환 (동적 범위 확장을 위해)
oa excel table-write --data-file "sales.csv" --table-name "SalesData" --table-style "TableStyleMedium5"

# 2단계: Excel Table 확인
oa excel table-list

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

### 6. 회계 데이터 정리 자동화
```bash
# 문자열 형식의 회계 데이터를 숫자로 일괄 변환
oa excel range-convert --range "A2:F1000" --expand table --remove-currency --remove-comma
# → "₩1,234,567" → 1234567로 자동 변환
# → 피벗테이블이나 계산에 바로 사용 가능

# 백분율 데이터 변환
oa excel range-convert --range "G1:G100" --parse-percent
# → "15.5%" → 0.155로 변환하여 수식에서 바로 활용

# 괄호형 음수 처리 (회계 양식)
oa excel range-convert --range "H1:H200" --remove-comma
# → "(1,000)" → -1000으로 자동 변환
```

### 7. 에러 방지 패턴
```bash
# 안전한 워크플로우: 확인 → 연결 → 작업
oa excel workbook-list | grep "target.xlsx"  # 파일 열림 확인
# 있으면: --workbook-name 사용, 없으면: --file-path로 열기
oa excel range-read --workbook-name "target.xlsx" --range "A1:C10"
```

## 🤖 AI 지원 데이터 분석 기능

### 스마트 데이터 패턴 감지 (data-analyze)
AI가 Excel 데이터를 자동으로 분석하여 피벗테이블 준비 상태를 평가합니다:

**🔍 자동 감지 패턴**:
- **교차표 형식**: 월/분기가 열로 배치된 형태 감지
- **다단계 헤더**: 중첩된 헤더 구조 인식
- **병합된 셀**: 빈 셀로 인한 데이터 불일치 탐지
- **소계 혼재**: 데이터와 소계가 섞여있는 패턴 분석
- **넓은 형식**: 여러 지표가 열로 나열된 구조 식별

**🎯 지능형 분석 결과**:
- 피벗테이블 준비도 0.0~1.0 점수로 평가
- 감지된 문제별 우선순위 권장사항 제공
- 다음 단계 명령어 자동 제안 (변환 타입 포함)

### 지능형 데이터 변환 (data-transform)
AI가 감지한 문제를 최적 순서로 자동 해결합니다:

**🔄 자동 변환 알고리즘 (auto 모드)**:
1. **병합셀 해제 우선**: 데이터 무결성 확보
2. **소계 제거**: 순수 데이터만 추출
3. **헤더 정리**: 다단계 헤더를 단일 헤더로 결합
4. **교차표 변환**: 피벗테이블용 세로 형식으로 변환

**📊 변환 결과 분석**:
- 변환 전후 데이터 크기 비교 (행/열 변화율)
- 적용된 변환 목록과 순서 보고
- 데이터 확장비/감소비 자동 계산

**💡 AI 활용 내부 구조**:
- **pandas 지능형 활용**: DataFrame 패턴 분석으로 데이터 구조 자동 인식
- **통계적 패턴 매칭**: 공통 Excel 문제 패턴을 학습된 알고리즘으로 탐지
- **컨텍스트 인식**: 헤더명, 데이터 분포, 빈 셀 패턴을 종합적으로 분석
- **최적화된 변환 순서**: 데이터 손실 최소화를 위한 단계별 변환 전략

## ✨ 특별 기능

- **🤖 AI 데이터 분석**: 5가지 일반적인 Excel 데이터 문제를 자동 감지하고 해결방안 제시
- **🔄 지능형 자동 변환**: AI가 감지한 모든 문제를 최적 순서로 자동 해결
- **📊 스마트 형식 변환**: 쉼표, 통화기호(₩,$,€,¥,£), 백분율, 괄호형 음수를 숫자로 자동 변환
- **자동 워크북 선택**: 옵션 없이 활성 워크북 자동 사용으로 Excel 재실행 없이 연속 작업
- **`--workbook-name`**: 파일명으로 직접 접근, 경로 불필요
- **워크북 연결 방법**: 옵션 없음(활성), `--file-path`(파일), `--workbook-name`(이름)
- **🎯 자동 배치**: 피벗테이블과 차트가 겹치지 않게 자동으로 빈 공간 찾아 배치
- **⚠️ 겹침 검사**: 지정된 위치의 충돌 여부를 사전 확인하여 경고 제공
- **JSON 최적화**: 모든 출력이 AI 에이전트 파싱에 최적화
- **한글 파일명 지원**: macOS에서 한글 자소분리 문제 자동 해결
- **39개 Excel 명령어**: 워크북/시트/데이터/차트/피벗/도형/슬라이서 전체 지원

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