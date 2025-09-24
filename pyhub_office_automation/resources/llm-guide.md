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

## 📈 LLM Context 최적화 전략 (대용량 데이터 처리)

### 🎯 문제 상황
- 대용량 Excel 테이블 (10만+ 행) 처리 시 LLM Context 한계 초과
- 전체 데이터를 무조건 로드하여 토큰 낭비
- AI 에이전트의 효율적인 데이터 분석 방해

### 🚀 해결책: table-read 고급 옵션 활용

#### 1️⃣ 페이징(Pagination) 전략
```bash
# 대용량 테이블을 작은 청크로 나누어 처리
oa excel table-read --table-name "BigData" --limit 100
oa excel table-read --table-name "BigData" --offset 100 --limit 100
oa excel table-read --table-name "BigData" --offset 200 --limit 100

# AI 워크플로우 예제
# 1단계: 첫 100행으로 데이터 구조 파악
oa excel table-read --table-name "Sales" --limit 100
# 2단계: 중간 구간 분석
oa excel table-read --table-name "Sales" --offset 5000 --limit 100
# 3단계: 최신 데이터 확인
oa excel table-read --table-name "Sales" --offset -100
```

#### 2️⃣ 지능형 샘플링 전략
```bash
# 균등 분포 샘플링 (첫 20% + 중간 60% + 마지막 20%)
oa excel table-read --table-name "BigData" --limit 50 --sample-mode

# 사용 시나리오:
# - 50만 행 테이블 → 50행 샘플링으로 전체 패턴 파악
# - 시계열 데이터 → 초기/중기/최근 패턴 모두 포함
# - 카테고리 데이터 → 다양한 카테고리 골고루 샘플링
```

#### 3️⃣ 컬럼 선택 전략 (Context 추가 절약)
```bash
# 분석에 필요한 컬럼만 선택
oa excel table-read --table-name "Sales" --columns "날짜,매출,지역" --limit 200

# 단계적 분석 워크플로우
# 1단계: 핵심 지표만 분석
oa excel table-read --table-name "Sales" --columns "매출,수익" --limit 100
# 2단계: 세부 분류 분석
oa excel table-read --table-name "Sales" --columns "제품,카테고리,매출" --limit 100
```

### 🧠 AI 에이전트 Context 관리 Best Practice

#### ✅ 권장 패턴: 점진적 데이터 탐색
```bash
# Step 1: 전체 구조 파악 (테이블 메타데이터)
oa excel table-list --format json
# → 테이블 크기, 컬럼 구조, 샘플 데이터 확인

# Step 2: 핵심 데이터만 추출 (Context 절약)
oa excel table-read --table-name "BigTable" --limit 50 --sample-mode
# → 전체 패턴 파악용 샘플 데이터

# Step 3: 특정 분석을 위한 타겟 데이터
oa excel table-read --table-name "BigTable" --columns "핵심컬럼1,핵심컬럼2" --limit 200
# → 분석 목적에 최적화된 데이터셋
```

#### ❌ 피해야 할 패턴: Context 낭비
```bash
# 🚫 전체 데이터 무조건 로드
oa excel table-read --table-name "BigTable"  # 10만 행 모두 로드

# 🚫 불필요한 컬럼 포함
oa excel table-read --table-name "Sales" --limit 1000  # 50개 컬럼 모두 포함

# 🚫 반복적인 대용량 읽기
# 매번 새로 대용량 데이터 읽지 말고 이전 결과 활용
```

### 📊 실제 사용 시나리오

#### 시나리오 1: 월간 보고서 생성 (50만 행 데이터)
```bash
# 전략: 시간대별 샘플링 + 핵심 지표 집중
oa excel table-read --table-name "MonthlyData" --columns "날짜,매출,고객수" --limit 100 --sample-mode
# → 월 전체 트렌드 파악

oa excel table-read --table-name "MonthlyData" --columns "제품,매출" --offset 0 --limit 500
# → 초기 제품별 상세 분석

oa excel table-read --table-name "MonthlyData" --columns "지역,매출" --offset -1000
# → 최신 지역별 실적 분석
```

#### 시나리오 2: 이상치 탐지 (100만 행 데이터)
```bash
# 전략: 구간별 샘플링으로 이상 패턴 탐지
oa excel table-read --table-name "Transactions" --columns "금액,시간,계정" --limit 200 --sample-mode
# → 전체 구간에서 이상치 패턴 확인

# 이상치 발견 시 해당 구간 집중 분석
oa excel table-read --table-name "Transactions" --offset 50000 --limit 1000
# → 특정 구간 상세 분석
```

#### 시나리오 3: 다차원 분석 (카테고리 × 시간 × 지역)
```bash
# 전략: 차원별 순차 분석
oa excel table-read --table-name "MultiDim" --columns "카테고리,매출" --limit 500
# → 카테고리 차원 분석

oa excel table-read --table-name "MultiDim" --columns "지역,매출" --limit 500
# → 지역 차원 분석

oa excel table-read --table-name "MultiDim" --columns "날짜,매출" --limit 300
# → 시간 차원 분석
```

### ⚡ 성능 최적화 효과

| 데이터 크기 | 기존 방식 | 최적화 방식 | Context 절약 |
|-------------|-----------|-------------|--------------|
| 10만 행 × 20열 | 200만 토큰 | **10만 토큰** | **-95%** |
| 50만 행 × 10열 | 500만 토큰 | **5만 토큰** | **-99%** |
| 컬럼 선택 (5/20열) | 100% | **25%** | **-75%** |

### 🎯 AI 에이전트 가이드라인

1. **항상 table-list로 시작**: 데이터 크기 사전 파악
2. **크기별 전략 수립**:
   - 1만 행 이하: 전체 로드 가능
   - 1만~10만 행: limit + 컬럼 선택
   - 10만 행 이상: sample-mode + 페이징 조합
3. **점진적 분석**: 샘플 → 타겟 → 상세 순서
4. **Context 예산 관리**: 토큰 사용량 지속 모니터링

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

# 🆕 혁신적 테이블 컨텍스트 즉시 파악 (AI 최적화)
oa excel table-list --format json
# → 한 번의 호출로 모든 테이블의 구조, 컬럼, 샘플 데이터 획득
# → 추가 API 호출 없이 완전한 비즈니스 컨텍스트 이해 가능

# 특정 워크북 상세 분석
oa excel workbook-info --workbook-name "Sales.xlsx" --include-sheets
```

### 1-1. 🔥 Enhanced Table-Driven Analysis (혁신적 워크플로우)

**기존 방식** (3번 호출 필요):
```bash
oa excel table-list              # 테이블 이름만 확인
oa excel range-read --range ...  # 구조 파악을 위한 추가 호출
oa excel range-read --range ...  # 샘플 데이터를 위한 추가 호출
```

**🆕 새로운 방식** (1번 호출로 완료):
```bash
oa excel table-list --format json
```

#### 📊 실제 응답 예제 (GameData 테이블)
```json
{
  "success": true,
  "command": "table-list",
  "data": {
    "tables": [
      {
        "name": "GameData",
        "sheet": "Data",
        "range": "$A$1:$K$999",
        "row_count": 999,
        "column_count": 11,
        "has_headers": true,
        "style": "TableStyleMedium3",
        "data_rows": 998,
        "columns": [
          "순위", "게임명", "플랫폼", "발행일", "장르", "퍼블리셔",
          "북미 판매량", "유럽 판매량", "일본 판매량", "기타 판매량", "글로벌 판매량"
        ],
        "sample_data": [
          ["1.0", "Wii 스포츠", "닌텐도 Wii", "2006.0", "스포츠", "닌텐도", "41.49", "29.02", "3.77", "8.46", "82.74"],
          ["2.0", "슈퍼 마리오브라더스", "닌텐도S", "1985.0", "횡스크롤", "닌텐도", "29.08", "3.58", "6.81", "0.77", "40.24"],
          ["3.0", "마리오 카트 Wii", "닌텐도 Wii", "2008.0", "레이싱", "닌텐도", "15.85", "12.88", "3.79", "3.31", "35.82"]
        ]
      }
    ],
    "summary": {
      "total_tables": 1,
      "sheets_with_tables": 1,
      "sheets_scanned": 3
    }
  }
}
```

#### 🚀 AI 에이전트 즉시 활용 시나리오

**🎯 시나리오 1: 즉시 데이터 분석 제안**
```bash
# 1번의 table-list 호출로 AI가 즉시 파악:
# - 게임 판매량 데이터 (998개 게임)
# - 지역별 판매량 컬럼 4개 (북미, 유럽, 일본, 기타)
# - 플랫폼/장르별 분류 가능
# - 시계열 데이터 (발행일)

# AI 즉시 제안 가능:
"998개 게임의 글로벌 판매 데이터를 발견했습니다!
- 지역별 판매량 비교 차트를 만들어드릴까요?
- 장르별/플랫폼별 분석은 어떠세요?
- 발행 연도별 트렌드 분석도 가능합니다!"
```

**📊 시나리오 2: 적절한 차트 타입 자동 추천**
```bash
# 컬럼 구조 분석으로 AI가 즉시 추천:
# - 순위 + 글로벌 판매량 → Top 10 막대형 차트
# - 지역별 판매량 → 지역 비교 차트
# - 장르별 집계 → 원형/도넛 차트
# - 발행일 × 판매량 → 시계열 선형 차트

# 자동 생성 예제:
oa excel chart-add --sheet "Data" --data-range "GameData[글로벌 판매량]" --chart-type "Column" --title "게임별 글로벌 판매량 Top 10"
```

**🔍 시나리오 3: 데이터 품질 즉시 평가**
```bash
# 샘플 데이터만으로 AI가 즉시 평가:
# - "41.49" 등 숫자 데이터 정상 (백만장 단위)
# - "Wii 스포츠" 등 게임명 정상
# - "2006.0" 발행일 형식 확인 필요
# - NULL 값 없음 확인

# AI 즉시 제안:
"데이터 품질이 양호합니다. 998행 모두 분석 가능한 상태로 보입니다.
발행일이 실수 형태인데 연도로 변환하여 분석하겠습니다."
```

#### ⚡ 성능 혁신: 67% API 호출 감소
| 작업 | 기존 방식 | 새로운 방식 | 개선율 |
|------|-----------|-------------|---------|
| 테이블 구조 파악 | 3번 호출 | **1번 호출** | **-67%** |
| 컬럼 정보 확인 | 별도 호출 필요 | **즉시 제공** | **-100%** |
| 샘플 데이터 확인 | 별도 호출 필요 | **즉시 제공** | **-100%** |
| 비즈니스 컨텍스트 | 수동 추론 | **자동 파악** | **즉시** |

#### 🤖 AI 에이전트 Best Practice

```bash
# ✅ 권장: 작업 시작 시 한 번만 호출
oa excel table-list --format json

# ✅ 모든 정보를 한 번에 받아서 분석/제안 진행
# - 테이블 구조 즉시 파악
# - 적절한 차트 타입 자동 추천
# - 데이터 품질 사전 검증
# - 비즈니스 시나리오 자동 제안

# ❌ 피해야 할 패턴: 불필요한 추가 호출
# oa excel range-read ... (이미 sample_data에 포함됨)
# oa excel workbook-info ... (이미 충분한 컨텍스트 제공됨)
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

### 2-1. 🆕 Table Metadata Management (테이블 메타데이터 관리)

```bash
# 🎯 개별 테이블 심화 분석 및 메타데이터 생성
oa excel table-analyze --table-name "GameData" --format json
# → 특정 테이블의 상세 메타데이터 자동 생성
# → 데이터 타입 추론, 태그 자동 생성
# → Metadata 시트에 체계적으로 저장

# 🔄 워크북 전체 테이블 일괄 메타데이터 생성
oa excel metadata-generate --format json
# → 모든 테이블 자동 분석
# → 룰 기반 비즈니스 컨텍스트 추론
# → 통합 메타데이터 관리 시스템

# 📊 메타데이터 포함 워크북 정보
oa excel workbook-info --include-metadata --format json
# → 기존 기능 + 메타데이터 통계
# → 관리되는 테이블 vs 미관리 테이블 현황
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

