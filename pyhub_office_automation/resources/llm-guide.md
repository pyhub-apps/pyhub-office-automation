# LLM/AI 에이전트를 위한 사용 지침

## 🎯 핵심 명령어

- `oa info`: 패키지 정보 확인
- `oa excel list`: Excel 명령어 목록
- `oa excel workbook-list`: 현재 열린 워크북 확인

## 🔗 연결 방법 (자동 선택 시스템)

**워크북 연결 우선순위**:
1. **--file-path**: 특정 파일 경로로 연결
2. **--workbook-name**: 열린 워크북 이름으로 연결
3. **옵션 없음**: 활성 워크북 자동 사용 (기본값)

**시트 연결 우선순위**:
1. **--sheet**: 시트 이름 지정
2. **범위 내 시트**: "Sheet1!A1:C10" 형식
3. **옵션 없음**: 활성 시트 자동 사용

## 🚀 AI 에이전트 추천 워크플로우

### 1. Context Discovery (상황 파악)

```bash
# 전체 환경 파악 - 모든 작업 전 필수
oa excel workbook-list --detailed --format json

# 활성 워크북 구조 분석
oa excel workbook-info --include-sheets --format json

# 특정 워크북 상세 분석
oa excel workbook-info --workbook-name "Sales.xlsx" --include-sheets
```

### 2. Safe Operation (안전한 작업)

```bash
# 패턴 1: 기존 워크북 확인 후 연결
oa excel workbook-list | grep "target.xlsx"
# 있으면: --workbook-name "target.xlsx" 사용
# 없으면: --file-path "/path/to/target.xlsx" 사용

# 패턴 2: 활성 워크북 직접 사용 (가장 안전)
oa excel range-read --range "A1:C10"
oa excel range-write --range "D1" --data '["결과"]'
```

### 3. Efficient Batch Processing (효율적 일괄 처리)

```bash
# 한 번 열고 여러 작업 수행
oa excel workbook-open --file-path "report.xlsx"
oa excel sheet-add --name "Analysis"
oa excel range-write --range "A1:C1" --data '["Name", "Value", "Status"]'
oa excel table-write --data-file "data.json" --table-name "Results"
oa excel chart-add --data-range "A1:C10" --chart-type "column" --auto-position
```

### 4. Multi-Workbook Operations (다중 워크북 작업)

```bash
# 워크북별 작업 분리
oa excel range-read --workbook-name "Source.xlsx" --range "A1:Z100" --output-file "temp_data.json"
oa excel range-write --workbook-name "Target.xlsx" --range "A1" --data-file "temp_data.json"

# 워크북 간 데이터 이동
oa excel table-read --workbook-name "Raw.xlsx" --range "Data!A1" --expand table --output-file "processed.csv"
oa excel table-write --workbook-name "Report.xlsx" --data-file "processed.csv" --table-name "Summary"
```

### 5. Error Prevention & Recovery (에러 방지 및 복구)

```bash
# 작업 전 검증
oa excel workbook-list --format json  # JSON 파싱으로 워크북 존재 확인
oa excel workbook-info --include-sheets  # 시트 구조 확인

# 실패 시 대안 경로
# 1차: 활성 워크북 사용 시도
oa excel range-read --range "A1:C10"
# 실패 시: 명시적 연결 시도
oa excel range-read --workbook-name "파일명.xlsx" --range "A1:C10"
```

## 💡 스마트 선택 가이드

### 언제 어떤 연결 방법을 사용할까?

**옵션 없음 (권장)**:
- ✅ 사용자가 Excel에서 작업 중인 파일
- ✅ 연속된 여러 작업
- ✅ 간단하고 빠른 작업

**--workbook-name**:
- ✅ 여러 워크북이 열려있을 때
- ✅ 특정 워크북 지정 필요
- ✅ 워크북 이름을 정확히 알고 있을 때

**--file-path**:
- ✅ 워크북이 아직 열려있지 않을 때
- ✅ 자동화 스크립트
- ✅ 절대적으로 특정 파일이 필요할 때

## ⚠️ 에러 방지 핵심 원칙

1. **항상 먼저 확인**: `workbook-list`로 현재 상태 파악
2. **JSON 활용**: `--format json`으로 파싱 가능한 출력 사용
3. **단계적 접근**: 간단한 연결부터 시도 (옵션 없음 → --workbook-name → --file-path)
4. **범위 검증**: 데이터 읽기 전 시트 구조 확인
5. **경로 정확성**: 절대 경로 사용 권장

## 🛡️ 보안 & 개인정보

- 문서 내용은 AI 학습에 사용되지 않음
- 로컬 처리만 수행, 외부 전송 없음
- 임시 파일 자동 삭제

