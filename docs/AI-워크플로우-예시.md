# 🤖 AI 에이전트 워크플로우 예시

이 문서는 사용자와 AI 에이전트 간의 실제 대화를 통한 업무 처리 시나리오를 제공합니다. AI 에이전트가 초심자부터 고급 사용자까지 어떻게 대응해야 하는지 구체적인 예시를 보여줍니다.

---

## 🔰 시나리오 1: 초심자 - 월간 보고서 자동화

### 상황
- **사용자**: Excel 기본 사용만 가능한 사무직 직원
- **목표**: 매달 만드는 판매 보고서를 자동화하고 싶음
- **현재 상태**: Excel에 판매 데이터가 입력되어 있음

### 대화 시나리오

**Human**: 매달 만드는 판매 보고서를 자동화하고 싶어

**AI**: 네, 월간 판매 보고서 자동화를 도와드리겠습니다! 먼저 현재 작업 환경을 확인해볼게요.

*[AI 내부 판단: 초심자 감지 신호 - "자동화하고 싶어" + 구체적 방법 없음]*

```bash
# AI 실행
oa excel workbook-list --detailed --format json
```

**AI**: 현재 "2024년_판매데이터.xlsx"가 열려있네요. 이 파일로 작업할까요?

시트도 확인해보겠습니다.

```bash
# AI 실행
oa excel workbook-info --include-sheets --format json
```

**AI**: 현재 "1월", "2월", "3월" 시트가 있고, "요약" 시트도 있네요.
보고서는 보통 어떤 내용을 포함하시나요? 예를 들어:
- 📊 월별 매출 차트
- 📈 제품별 판매 현황
- 🏆 베스트 셀러 순위
- 📋 요약 통계표

**Human**: 월별 매출 차트랑 제품별 순위 정도면 될 것 같아

**AI**: 좋아요! 그럼 다음 순서로 보고서를 만들어드리겠습니다:

1️⃣ 먼저 데이터를 확인하고 정리하겠습니다
2️⃣ 월별 매출 차트를 만들겠습니다
3️⃣ 제품별 순위 표를 만들겠습니다
4️⃣ 보고서 시트에 정리해드리겠습니다

첫 번째 단계부터 시작할게요!

```bash
# AI 실행
oa excel range-read --sheet "1월" --range "A1:D100" --format json
```

**AI**: 1월 데이터를 확인했어요. "제품명", "매출액", "수량", "날짜" 컬럼이 있네요.
다른 월 데이터도 같은 형식인지 확인해보겠습니다.

```bash
# AI 실행
oa excel range-read --sheet "2월" --range "A1:D100" --format json
oa excel range-read --sheet "3월" --range "A1:D100" --format json
```

**AI**: 완벽해요! 모든 시트가 같은 형식이네요. 이제 전체 데이터를 통합해서 분석용 시트를 만들겠습니다.

```bash
# AI 실행
oa excel sheet-add --name "분석데이터"
```

**AI**: 이제 각 월 데이터를 하나로 합치겠습니다...

```bash
# AI 실행
oa excel table-read --sheet "1월" --range "A1" --expand table --output-file "jan_data.csv"
oa excel table-read --sheet "2월" --range "A1" --expand table --output-file "feb_data.csv"
oa excel table-read --sheet "3월" --range "A1" --expand table --output-file "mar_data.csv"
```

*[월별 데이터를 순차적으로 통합 시트에 추가]*

```bash
# AI 실행
oa excel table-write --sheet "분석데이터" --range "A1" --data-file "jan_data.csv"
```

**AI**: 데이터 통합이 완료되었습니다! 이제 월별 매출 차트를 만들어보겠습니다.

```bash
# AI 실행
oa excel chart-add --sheet "분석데이터" --data-range "A1:C50" --chart-type "column" --title "월별 매출 현황" --auto-position
```

**AI**: 차트가 생성되었어요! 다음으로 제품별 순위 표를 만들겠습니다. 피벗테이블을 사용해서 자동으로 정리해드릴게요.

```bash
# AI 실행
oa excel pivot-create --source-range "분석데이터!A1" --expand table --dest-sheet "분석데이터" --dest-range "F1"
oa excel pivot-configure --pivot-name "PivotTable1" --row-fields "제품명" --value-fields "매출액:Sum" --clear-existing
```

**AI**: 피벗테이블이 완성되었어요! 이제 마지막으로 보고서 시트를 만들어서 모든 내용을 정리해드리겠습니다.

```bash
# AI 실행
oa excel sheet-add --name "월간보고서"
oa excel range-write --sheet "월간보고서" --range "A1" --data "🏢 월간 판매 보고서"
```

**Human**: 와 이렇게 쉽게 되는구나! 다음달에도 똑같이 할 수 있어?

**AI**: 네, 물론이에요! 다음달에는 더 간단합니다:

1. 새로운 월 데이터를 시트에 추가
2. "분석데이터" 시트에서 범위만 늘려주기
3. 피벗테이블 새로고침 한 번이면 끝!

이렇게 말씀해주시면 됩니다: "4월 데이터 추가해서 보고서 업데이트해줘"

---

## 🔧 시나리오 2: 중급자 - 여러 파일 데이터 통합 분석

### 상황
- **사용자**: Excel 중급 수준, 피벗테이블과 차트 경험 있음
- **목표**: 지점별로 나누어진 파일들을 통합해서 전체 분석
- **현재 상태**: 3개 지점 파일이 각각 열려있음

### 대화 시나리오

**Human**: 서울, 부산, 대구 지점 파일들을 합쳐서 전국 현황을 분석하고 싶어. 지점별 비교도 하고 싶고

**AI**: 네, 지점별 파일 통합 분석을 도와드리겠습니다. 먼저 현재 열린 파일들을 확인해볼게요.

*[AI 내부 판단: 중급자 - 구체적 목표, 기술 용어 사용]*

```bash
# AI 실행
oa excel workbook-list --detailed --format json
```

**AI**: 확인했습니다!
- 서울지점_2024.xlsx
- 부산지점_2024.xlsx
- 대구지점_2024.xlsx

모든 파일이 열려있네요. 각 파일의 데이터 구조를 확인해보겠습니다.

```bash
# AI 실행
oa excel workbook-info --workbook-name "서울지점_2024.xlsx" --include-sheets
oa excel workbook-info --workbook-name "부산지점_2024.xlsx" --include-sheets
oa excel workbook-info --workbook-name "대구지점_2024.xlsx" --include-sheets
```

**AI**: 모든 파일이 동일한 구조네요. "판매데이터" 시트에 같은 컬럼들이 있습니다. 통합 분석용 새 파일을 만들까요?

```bash
# AI 실행
oa excel workbook-create --name "전국통합분석" --save-path "전국통합분석_2024.xlsx"
```

**AI**: 이제 각 지점 데이터를 읽어서 통합하겠습니다. 지점 구분을 위해 "지점" 컬럼을 추가하면서 진행할게요.

```bash
# AI 실행
oa excel table-read --workbook-name "서울지점_2024.xlsx" --sheet "판매데이터" --range "A1" --expand table --output-file "seoul_data.csv"
oa excel table-read --workbook-name "부산지점_2024.xlsx" --sheet "판매데이터" --range "A1" --expand table --output-file "busan_data.csv"
oa excel table-read --workbook-name "대구지점_2024.xlsx" --sheet "판매데이터" --range "A1" --expand table --output-file "daegu_data.csv"
```

**AI**: 데이터를 모두 읽었습니다. 이제 통합 시트에 지점별로 구분해서 추가하겠습니다.

```bash
# AI 실행
oa excel sheet-add --name "통합데이터"

# 헤더와 함께 서울 데이터 먼저 추가
oa excel table-write --sheet "통합데이터" --range "A1" --data-file "seoul_data.csv"

# 지점 컬럼 추가 (서울)
oa excel range-write --sheet "통합데이터" --range "F1" --data "지점"
oa excel range-write --sheet "통합데이터" --range "F2:F1000" --data "서울"
```

**Human**: 피벗테이블로 지점별 비교 분석하고, 차트도 지점별로 나누어서 보여줘

**AI**: 네, 피벗테이블과 차트를 만들어드리겠습니다. 먼저 전체 데이터를 Excel Table로 만들어서 동적 범위로 설정하겠습니다.

*[AI 내부 판단: 사용자가 피벗테이블 언급 - 중급자 확실, 빠르게 진행]*

```bash
# AI 실행
oa excel table-create --sheet "통합데이터" --range "A1:F1000" --table-name "전국판매데이터" --table-style "TableStyleMedium5"
```

**AI**: 이제 분석용 시트를 만들고 피벗테이블을 생성하겠습니다.

```bash
# AI 실행
oa excel sheet-add --name "지점별분석"

# 피벗테이블 생성
oa excel pivot-create --source-range "전국판매데이터" --dest-sheet "지점별분석" --dest-range "A1"

# 피벗테이블 설정
oa excel pivot-configure --pivot-name "PivotTable1" \
  --row-fields "지점,제품카테고리" \
  --value-fields "매출액:Sum,수량:Sum" \
  --column-fields "월"
```

**AI**: 피벗테이블이 완성되었습니다! 이제 지점별 비교 차트를 만들겠습니다.

```bash
# AI 실행
# 지점별 총 매출 차트
oa excel chart-add --sheet "지점별분석" --data-range "A1:D4" --chart-type "column" --title "지점별 총 매출 비교" --position "F1"

# 월별 추세 차트 (지점별)
oa excel chart-add --sheet "지점별분석" --data-range "A1:M6" --chart-type "line" --title "지점별 월별 매출 추세" --auto-position --preferred-position "bottom"
```

**Human**: 대시보드 형태로 한 화면에서 모든 분석을 볼 수 있게 만들어줘

**AI**: 대시보드를 만들어드리겠습니다!

```bash
# AI 실행
oa excel sheet-add --name "전국판매대시보드"

# 대시보드용 요약 데이터 생성
oa excel pivot-create --source-range "전국판매데이터" --dest-sheet "전국판매대시보드" --dest-range "A25" --pivot-name "DashboardPivot1"
oa excel pivot-configure --pivot-name "DashboardPivot1" --row-fields "지점" --value-fields "매출액:Sum,수량:Sum"

# 여러 차트를 배치
oa excel chart-add --sheet "전국판매대시보드" --data-range "A25:C28" --chart-type "column" --title "지점별 매출" --position "A1" --width 300 --height 200

oa excel chart-add --sheet "전국판매대시보드" --data-range "A25:B28" --chart-type "pie" --title "지점별 매출 비중" --position "E1" --width 300 --height 200

# 슬라이서 추가 (인터랙티브 필터)
oa excel slicer-add --pivot-name "DashboardPivot1" --field "지점" --position "A15"
```

**AI**: 대시보드가 완성되었습니다! 슬라이서를 클릭하면 특정 지점만 볼 수 있고, 차트들이 실시간으로 업데이트됩니다.

---

## 🚀 시나리오 3: 고급자 - 복잡한 대시보드 구축

### 상황
- **사용자**: Excel 고급 사용자, 매크로와 복잡한 공식 사용 경험
- **목표**: 경영진용 종합 대시보드 구축 (매출, 수익성, 예측 포함)
- **현재 상태**: 여러 시스템에서 추출한 다양한 데이터 파일들

### 대화 시나리오

**Human**: 경영진 보고용 대시보드를 만들려고 해. 매출 추세, 수익성 분석, 예측 모델까지 포함해서. 데이터는 ERP, 회계시스템, 외부 시장 데이터 등 여러 소스에서 가져온 거야.

**AI**: 경영진용 종합 대시보드 구축을 도와드리겠습니다. 상당히 복잡한 프로젝트네요!

먼저 현재 데이터 소스들을 파악해보겠습니다.

*[AI 내부 판단: 고급 사용자 - 복잡한 요구사항, 기술적 배경 있음]*

```bash
# AI 실행
oa excel workbook-list --detailed --format json
```

**AI**: 현재 이런 파일들이 열려있네요:
- ERP_매출데이터_Q4.xlsx
- 회계_손익분석_2024.xlsx
- 시장데이터_업계현황.xlsx
- 예산계획_2024.xlsx

각 데이터의 구조와 연결점을 분석해보겠습니다.

```bash
# AI 실행
oa excel workbook-info --workbook-name "ERP_매출데이터_Q4.xlsx" --include-sheets --include-properties
oa excel workbook-info --workbook-name "회계_손익분석_2024.xlsx" --include-sheets --include-properties
```

**AI**: 데이터 스키마를 분석해보니 다음과 같은 연결 구조가 가능해 보입니다:
- 공통 Key: 제품코드, 날짜
- 매출 (ERP) + 원가 (회계) = 수익성 분석
- 시장데이터로 벤치마킹 가능

단계적으로 구축해보겠습니다:

1️⃣ 데이터 통합 및 정규화
2️⃣ 핵심 KPI 계산
3️⃣ 예측 모델 구현
4️⃣ 인터랙티브 대시보드 구성

**Human**: 좋아. 데이터 품질 검증도 포함해줘. 그리고 드릴다운 기능도 있으면 좋겠어.

**AI**: 네, 데이터 품질 검증과 드릴다운 기능까지 포함해서 구축하겠습니다.

먼저 마스터 대시보드 파일을 생성하겠습니다.

```bash
# AI 실행
oa excel workbook-create --name "경영진_종합대시보드" --save-path "executive_dashboard_2024.xlsx"
```

**1단계: 데이터 통합 및 품질 검증**

```bash
# AI 실행
# 데이터 품질 검증 시트 생성
oa excel sheet-add --name "데이터품질검증"

# 각 소스별 데이터 검증
oa excel data-analyze --workbook-name "ERP_매출데이터_Q4.xlsx" --sheet "매출" --output-file "erp_analysis.json"
oa excel data-analyze --workbook-name "회계_손익분석_2024.xlsx" --sheet "손익" --output-file "accounting_analysis.json"
```

**AI**: 데이터 품질 분석 결과:
- ERP 데이터: 95% 완전성, 3개 중복 레코드 발견
- 회계 데이터: 98% 완전성, 날짜 형식 불일치 발견

품질 이슈를 먼저 해결하겠습니다.

```bash
# AI 실행
# 데이터 정리 및 표준화
oa excel sheet-add --name "정제된데이터"

# ERP 데이터 정리
oa excel table-read --workbook-name "ERP_매출데이터_Q4.xlsx" --sheet "매출" --range "A1" --expand table --output-file "erp_raw.csv"
oa excel data-transform --input-file "erp_raw.csv" --remove-duplicates --standardize-dates --output-file "erp_clean.csv"

# 정리된 데이터를 마스터 시트에 로드
oa excel table-write --sheet "정제된데이터" --range "A1" --data-file "erp_clean.csv" --table-name "CleanERP"
```

**2단계: 핵심 KPI 계산**

```bash
# AI 실행
oa excel sheet-add --name "KPI계산"

# 수익성 지표 계산 (매출 - 원가)
oa excel range-write --sheet "KPI계산" --range "A1:E1" --data '["기간", "매출", "원가", "수익", "수익률"]'

# 복합 공식으로 KPI 계산
oa excel range-write --sheet "KPI계산" --range "A2" --data "=정제된데이터.CleanERP[날짜]"
oa excel range-write --sheet "KPI계산" --range "E2" --data "=D2/B2*100"
```

**3단계: 예측 모델 구현**

```bash
# AI 실행
oa excel sheet-add --name "예측분석"

# 선형 회귀 기반 매출 예측
oa excel chart-add --sheet "예측분석" --data-range "KPI계산!A1:B20" --chart-type "scatter" --title "매출 추세 및 예측"
oa excel chart-configure --chart-name "Chart1" --add-trendline --trendline-type "linear" --extend-forecast 3
```

**4단계: 드릴다운 대시보드 구성**

```bash
# AI 실행
oa excel sheet-add --name "경영진대시보드"

# 최상위 KPI 대시보드
oa excel range-write --sheet "경영진대시보드" --range "A1" --data "📊 경영진 종합 대시보드"

# 여러 레벨의 피벗테이블 생성 (드릴다운용)
oa excel pivot-create --source-range "CleanERP" --dest-sheet "경영진대시보드" --dest-range "A20" --pivot-name "ExecutivePivot"
oa excel pivot-configure --pivot-name "ExecutivePivot" \
  --row-fields "사업부,제품군,제품" \
  --value-fields "매출:Sum,수익:Sum" \
  --column-fields "분기"

# 인터랙티브 슬라이서 추가
oa excel slicer-add --pivot-name "ExecutivePivot" --field "사업부" --position "A5"
oa excel slicer-add --pivot-name "ExecutivePivot" --field "분기" --position "E5"

# 핵심 차트들 배치
oa excel chart-pivot-create --pivot-name "ExecutivePivot" --chart-type "column" --title "사업부별 실적" --position "A30"
oa excel chart-add --sheet "경영진대시보드" --data-range "KPI계산!A1:E12" --chart-type "line" --title "수익률 추이" --auto-position
```

**Human**: 실시간 업데이트는 어떻게 하지? 그리고 알람 기능도 넣을 수 있어?

**AI**: 실시간 업데이트와 알람 기능을 구현해보겠습니다.

**실시간 업데이트 메커니즘:**

```bash
# AI 실행
# 데이터 연결 설정 (외부 데이터 소스)
oa excel sheet-add --name "데이터연결설정"

# 자동 새로고침 스크립트 생성 시트
oa excel range-write --sheet "데이터연결설정" --range "A1" --data "자동 새로고침 버튼"

# 매크로 대신 피벗테이블 새로고침 버튼 생성
oa excel range-write --sheet "경영진대시보드" --range "H1" --data "데이터 새로고침"
```

**알람 기능 (조건부 서식 기반):**

```bash
# AI 실행
oa excel sheet-add --name "알람및모니터링"

# KPI 임계값 설정
oa excel range-write --sheet "알람및모니터링" --range "A1:C5" --data '[
  ["지표", "임계값", "현재값"],
  ["월매출", "10000000", "=SUM(KPI계산!B:B)"],
  ["수익률", "15", "=AVERAGE(KPI계산!E:E)"],
  ["신규고객", "100", "=COUNT(정제된데이터!A:A)"],
  ["재고회전율", "6", "=12/AVERAGE(H:H)"]
]'

# 상태 표시 컬럼 추가 (조건부 공식)
oa excel range-write --sheet "알람및모니터링" --range "D1" --data "상태"
oa excel range-write --sheet "알람및모니터링" --range "D2" --data '=IF(C2>=B2,"✅ 목표달성","⚠️ 주의필요")'
```

**Human**: 완벽해! 이제 이 대시보드를 기반으로 주간 보고서도 자동 생성할 수 있어?

**AI**: 물론입니다! 주간 보고서 자동 생성 기능을 추가하겠습니다.

```bash
# AI 실행
oa excel sheet-add --name "주간보고서템플릿"

# 보고서 헤더
oa excel range-write --sheet "주간보고서템플릿" --range "A1" --data "📈 주간 경영 보고서"
oa excel range-write --sheet "주간보고서템플릿" --range "A3" --data '=TEXT(TODAY(),"yyyy년 mm월 dd일 기준")'

# 자동 요약 섹션
oa excel range-write --sheet "주간보고서템플릿" --range "A5:B10" --data '[
  ["📊 핵심 지표", ""],
  ["주간 매출", "=SUM(KPI계산!B:B)/4"],
  ["전주 대비", "=B6/B6*1.05-1"],
  ["목표 달성률", "=B6/10000000*100"],
  ["주요 이슈", "=IF(알람및모니터링!D2<>\"✅ 목표달성\",\"목표 미달성\",\"정상 운영\")"],
  ["", ""]
]'

# 차트 자동 복사 (링크)
oa excel chart-add --sheet "주간보고서템플릿" --data-range "KPI계산!A1:E7" --chart-type "line" --title "주간 실적 추이" --position "D5"
```

**AI**: 주간 보고서 자동 생성이 완료되었습니다!

**앞으로 사용 방법:**
1. **일일 업데이트**: `oa excel pivot-refresh --all` 실행
2. **주간 보고서 생성**: "주간보고서템플릿" 시트를 복사해서 새 파일로 저장
3. **알람 확인**: "알람및모니터링" 시트에서 상태 컬럼 확인

모든 차트와 표가 자동으로 최신 데이터를 반영하므로, 데이터만 업데이트하면 보고서가 자동 완성됩니다!

---

## 🔄 시나리오 4: 에러 대응 - 파일 연결 문제

### 상황
- **사용자**: 초심자, 기술적 문제에 당황함
- **문제**: 파일이 열리지 않고 에러 메시지가 계속 나타남

### 대화 시나리오

**Human**: 도와줘! 파일이 안 열리고 계속 에러가 나

**AI**: 괜찮아요, 차근차근 해결해보겠습니다! 먼저 현재 상황을 확인해볼게요.

*[AI 내부 판단: 초심자 + 에러 상황 - 진단부터 시작]*

```bash
# AI 실행
oa excel workbook-list --format json
```

**AI**: 현재 열린 Excel 파일이 없네요. 어떤 파일을 열려고 하셨나요? 파일 이름을 알려주세요.

**Human**: "월간데이터.xlsx" 파일인데, 어제까지는 잘 됐거든

**AI**: 아, "월간데이터.xlsx" 파일이군요. 몇 가지 방법으로 해결해보겠습니다.

먼저 파일이 어디에 있는지 확인하고 열어보겠습니다.

```bash
# AI 실행
oa excel workbook-open --file-path "월간데이터.xlsx"
```

*[에러 발생 가정]*

**AI**: 파일을 찾을 수 없다고 나오네요. 파일이 다른 위치에 있을 수 있어요.

이렇게 해보겠습니다:
1. 바탕화면에 있는지 확인
2. 다운로드 폴더에 있는지 확인
3. 문서 폴더에 있는지 확인

파일이 보통 어디에 있나요?

**Human**: 바탕화면에 있어야 하는데...

**AI**: 바탕화면 경로로 다시 시도해보겠습니다.

```bash
# AI 실행
oa excel workbook-open --file-path "C:\Users\%USERNAME%\Desktop\월간데이터.xlsx"
```

*[여전히 에러 가정]*

**AI**: 여전히 파일을 찾을 수 없네요. 혹시 파일명이 조금 다를 수도 있어요.

**이렇게 해볼까요?**
1. 바탕화면에서 "월간" 이라는 단어가 들어간 Excel 파일을 찾아보세요
2. 파일명을 정확히 복사해서 알려주세요
3. 아니면 새로운 파일로 작업을 시작할 수도 있어요

어떻게 하시겠어요?

**Human**: 파일명이 "월간_판매데이터_2024.xlsx" 였네

**AI**: 아, 그렇군요! 정확한 파일명으로 다시 열어보겠습니다.

```bash
# AI 실행
oa excel workbook-open --file-path "C:\Users\%USERNAME%\Desktop\월간_판매데이터_2024.xlsx"
```

**AI**: 성공했어요! 파일이 열렸습니다. 파일명이 조금 달랐던 것 같네요.

이제 어떤 작업을 도와드릴까요?

```bash
# AI 실행
oa excel workbook-info --include-sheets
```

**AI**: 현재 "1월", "2월", "3월", "요약" 시트가 있네요. 잘 열렸습니다!

**💡 앞으로 이런 문제를 피하는 팁:**
- 파일명에 특수문자나 공백이 많으면 문제가 생길 수 있어요
- 파일을 항상 같은 위치에 저장하면 찾기 쉬워요
- 다음번에는 "바탕화면의 월간 파일 열어줘" 라고 말씀해주시면 제가 찾아드릴게요!

---

## 🎯 워크플로우 패턴 분석

### 초심자 대응 패턴
1. **친근한 인사** → 불안감 해소
2. **목적 구체화** → 막연한 요청을 구체적으로
3. **현황 파악** → 작업 환경 확인
4. **단계별 설명** → 각 과정을 이해할 수 있게
5. **결과 해석** → 기술적 결과를 일반 언어로
6. **다음 단계 안내** → 연속성 있는 도움

### 중급자 대응 패턴
1. **빠른 상황 파악** → 기본 확인 후 본격 작업
2. **기술적 용어 사용** → 사용자 수준에 맞춤
3. **효율적 진행** → 불필요한 설명 생략
4. **고급 기능 활용** → 피벗테이블, 차트 등
5. **결과 중심 소통** → 과정보다는 결과에 집중

### 고급자 대응 패턴
1. **복잡한 요구사항 분석** → 전체 아키텍처 설계
2. **다단계 프로젝트 관리** → 체계적 접근
3. **고급 기능 조합** → 여러 도구 통합 활용
4. **확장성 고려** → 미래 요구사항까지 대비
5. **기술적 세부사항 제공** → 구현 방법까지 상세히

### 에러 대응 패턴
1. **안심시키기** → 문제 해결 가능함을 강조
2. **체계적 진단** → 단계별 원인 파악
3. **다중 해결책** → 여러 방법 시도
4. **재발 방지** → 예방 팁 제공
5. **대안 제시** → 원래 방법 안 될 때 다른 방법

이러한 워크플로우 패턴을 참고하여 사용자의 수준과 상황에 맞는 대응을 해주세요.