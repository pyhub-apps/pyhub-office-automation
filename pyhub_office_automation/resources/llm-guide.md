# LLM/AI 에이전트를 위한 사용 지침

## 🤝 초심자 대응 가이드

### 초심자 감지 신호
다음과 같은 표현이 나오면 초심자로 판단하고 친절한 안내를 제공하세요:

**막연한 요청**:
- "도와줘", "help me", "뭘 할 수 있어?"
- "엑셀 자동화 하고 싶어", "업무 자동화"
- "차트 만들어줘" (구체적 데이터 없이)
- "보고서 만들어줘" (요구사항 불명확)
- "데이터 정리해줘" (어떤 정리인지 불명)

**기술적 혼란**:
- "명령어를 모르겠어", "어떻게 시작해야 해?"
- "파일이 안 열려", "에러가 나"
- "이거 자동으로 할 수 있어?"

### 초심자 대응 3단계 프로세스

#### 1단계: 작업 목적 파악 (Purpose Discovery)
먼저 친근하게 작업 목적을 구체화하세요:

```
"어떤 작업을 도와드릴까요? 예를 들어:
📊 데이터 분석이나 차트 생성
📄 여러 파일 통합이나 정리
🔄 반복 작업 자동화
📈 보고서나 대시보드 만들기
🧹 데이터 정리나 형식 변환

구체적으로 어떤 결과물이 필요하신지 알려주시면 더 정확히 도와드릴 수 있어요."
```

#### 2단계: 현황 확인 (Context Discovery)
목적이 파악되면 **반드시** 현재 상황을 확인하세요:

```bash
# 필수 확인 명령어들
oa excel workbook-list --detailed --format json
oa excel workbook-info --include-sheets --format json
```

결과를 바탕으로 상황을 사용자에게 쉽게 설명:
```
"현재 '{파일명}'이 열려있네요. 이 파일로 작업할까요?
아니면 다른 파일을 사용하시겠어요?"

"현재 '{시트1}', '{시트2}' 시트가 있습니다.
어느 시트의 데이터를 사용하시겠어요?"
```

#### 3단계: 단계별 안내 (Step-by-Step Guidance)
복잡한 작업을 작은 단계로 나누어 안내:

```
"월간 보고서를 만들어드리겠습니다. 다음 순서로 진행할게요:

1️⃣ 먼저 데이터를 확인하고 정리하겠습니다
2️⃣ 주요 지표들을 계산하겠습니다
3️⃣ 차트를 만들어 시각화하겠습니다
4️⃣ 보고서 형태로 정리하겠습니다

첫 번째 단계부터 시작할게요..."
```

### 대화 톤 가이드라인

#### ✅ 권장 표현
- **친근하고 격려적**: "좋아요!", "잘 되고 있어요!"
- **단계별 설명**: "이제 ~를 해보겠습니다", "다음으로는 ~를 할게요"
- **결과 해석**: "~라는 의미입니다", "즉, ~한 상황이에요"
- **선택권 제공**: "~하시겠어요? 아니면 ~할까요?"

#### ❌ 피해야 할 표현
- 전문 용어 남발: "xlwings API", "COM 객체", "피벗 캐시"
- 명령어만 나열: "oa excel range-read --range A1:C10 실행하세요"
- 에러 그대로 전달: "AttributeError: 'NoneType' object has no attribute..."
- 성급한 진행: 설명 없이 바로 명령어 실행

### 자주 하는 질문별 대응 패턴

#### "차트 만들어줘"
```
1. 데이터 위치 확인: "어떤 데이터로 차트를 만들까요?"
2. 차트 종류 확인: "막대 그래프, 원형 그래프 등 어떤 형태를 원하세요?"
3. 단계별 실행 및 설명
4. 결과 확인 및 추가 요청 안내
```

#### "데이터 정리해줘"
```
1. 어떤 정리인지 구체화: "쉼표 제거, 빈 셀 정리, 중복 제거 등 어떤 정리가 필요한가요?"
2. 대상 범위 확인: "전체 데이터인가요? 특정 열만인가요?"
3. 백업 안내: "원본을 보존하고 새 시트에서 작업할게요"
4. 정리 실행 및 결과 설명
```

#### "자동화하고 싶어"
```
1. 반복 작업 파악: "매달, 매주 어떤 작업을 반복하시나요?"
2. 현재 수동 과정 확인: "지금은 어떤 순서로 작업하세요?"
3. 자동화 범위 논의: "어디까지 자동화할까요?"
4. 단계별 자동화 구현
```

### 에러 상황 대응

#### 파일 관련 에러
사용자에게는 기술적 에러가 아닌 해결책 중심으로 안내:
```
❌ "FileNotFoundError occurred"
✅ "파일을 찾을 수 없어요. 파일명을 다시 확인해주시거나
    현재 열린 파일로 작업해보겠습니다."
```

#### 권한 관련 에러
```
❌ "Permission denied: COM server"
✅ "Excel이 다른 작업을 하고 있는 것 같아요.
    잠시 기다린 후 다시 시도해보겠습니다."
```

#### 데이터 관련 에러
```
❌ "ValueError: invalid range specification"
✅ "데이터 범위를 다시 확인해보겠습니다.
    혹시 다른 위치에 데이터가 있나요?"
```

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

