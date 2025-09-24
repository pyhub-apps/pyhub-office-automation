# 🤖 AI 에이전트를 위한 사용자 요청 예시

이 문서는 사용자가 AI 에이전트에게 하는 일반적인 요청들과 그에 대응하는 `oa` 명령어 시퀀스를 정리한 것입니다.

## 📊 데이터 분석 요청

### 차트 생성 관련

#### 📈 "월별 매출 데이터를 차트로 만들어줘"
**사용자 표현 변형**:
- "매출 그래프 그려줘"
- "월별 추세를 보여줘"
- "판매 데이터 시각화해줘"

**AI 실행 시퀀스**:
```bash
# 1. 현황 파악
oa excel workbook-list --detailed
oa excel workbook-info --include-sheets

# 2. 데이터 확인
oa excel range-read --range "A1:B12" --sheet "매출데이터"

# 3. 차트 생성
oa excel chart-add --data-range "A1:B12" --chart-type "line" --title "월별 매출 추세" --auto-position
```

#### 🥧 "부서별 비중을 원형 그래프로 보여줘"
**사용자 표현 변형**:
- "파이 차트 만들어줘"
- "비율을 원그래프로"
- "점유율 차트"

**AI 실행 시퀀스**:
```bash
# 1. 데이터 위치 확인
oa excel range-read --range "A1:B10" --sheet "부서별데이터"

# 2. 원형 차트 생성
oa excel chart-add --data-range "A1:B10" --chart-type "pie" --title "부서별 점유율" --show-data-labels --auto-position
```

### 통계 및 요약 관련

#### 📊 "평균이랑 합계 구해줘"
**사용자 표현 변형**:
- "총합 계산해줘"
- "평균값 알려줘"
- "통계 내줘"

**AI 실행 시퀀스**:
```bash
# 1. 데이터 범위 확인
oa excel range-read --range "A1:C100" --expand table

# 2. 새 시트에 요약 생성
oa excel sheet-add --name "요약통계"

# 3. 통계 계산 (Excel 함수 사용)
oa excel range-write --sheet "요약통계" --range "A1:B5" --data '[
  ["항목", "값"],
  ["합계", "=SUM(Sheet1!C:C)"],
  ["평균", "=AVERAGE(Sheet1!C:C)"],
  ["최대값", "=MAX(Sheet1!C:C)"],
  ["최소값", "=MIN(Sheet1!C:C)"]
]'
```

#### 🔝 "상위 10개 제품 찾아서 정리해줘"
**사용자 표현 변형**:
- "베스트 10 뽑아줘"
- "많이 팔린 순서대로"
- "톱 랭킹 보여줘"

**AI 실행 시퀀스**:
```bash
# 1. 전체 데이터 읽기
oa excel table-read --range "A1" --expand table --output-file "temp_data.csv"

# 2. 정렬된 테이블 생성
oa excel sheet-add --name "상위10개"
oa excel table-write --sheet "상위10개" --range "A1" --data-file "temp_data.csv"

# 3. 정렬 적용
oa excel table-sort --sheet "상위10개" --table-name "테이블1" --sort-key "매출액" --order "desc"

# 4. 상위 10개만 필터링 (수동으로 범위 제한)
oa excel range-read --sheet "상위10개" --range "A1:D11"
```

## 🧹 데이터 정리 요청

### 형식 변환 관련

#### 💰 "쉼표랑 원화 기호 제거해줘"
**사용자 표현 변형**:
- "숫자로 바꿔줘"
- "계산할 수 있게 정리해줘"
- "문자를 숫자로 변환"

**AI 실행 시퀀스**:
```bash
# 1. 현재 데이터 확인
oa excel range-read --range "A1:Z100"

# 2. 형식 변환 (원본 보존)
oa excel sheet-add --name "정리된데이터"
oa excel range-read --range "A1:Z100" --output-file "temp_data.json"
oa excel range-write --sheet "정리된데이터" --range "A1" --data-file "temp_data.json"

# 3. 숫자 형식 변환
oa excel range-convert --sheet "정리된데이터" --range "A1:Z100" --remove-comma --remove-currency
```

#### 📊 "백분율을 소수로 바꿔줘"
**사용자 표현 변형**:
- "퍼센트를 숫자로"
- "50%를 0.5로"
- "비율 계산용으로 변환"

**AI 실행 시퀀스**:
```bash
# 백분율 변환
oa excel range-convert --range "D1:D100" --parse-percent --expand down
```

### 정리 및 정렬 관련

#### 🗂️ "중복 제거하고 정렬해줘"
**사용자 표현 변형**:
- "같은 것들 정리해줘"
- "깔끔하게 정리"
- "중복 없이 순서대로"

**AI 실행 시퀀스**:
```bash
# 1. Excel Table로 변환 (중복 제거 기능 활용)
oa excel table-create --range "A1:C100" --table-name "정리테이블"

# 2. 정렬 적용
oa excel table-sort --table-name "정리테이블" --sort-key "이름" --order "asc"

# 3. 결과 확인
oa excel table-list
```

#### 📋 "빈 셀이나 빈 행 제거해줘"
**사용자 표현 변형**:
- "공백 정리해줘"
- "비어있는 곳 제거"
- "빈 부분 정리"

**AI 실행 시퀀스**:
```bash
# 1. 데이터 범위 확인
oa excel range-read --range "A1" --expand table

# 2. 새 시트에 정리된 데이터 생성
oa excel sheet-add --name "정리완료"

# 3. 빈 셀 제외하고 복사 (Excel 기능 활용)
# 실제로는 Go To Special 등의 Excel 기능을 사용해야 하므로
# 이 경우 사용자에게 수동 작업을 안내하거나 대안 제시
```

## 📄 보고서 작성 요청

### 데이터 통합 관련

#### 📊 "여러 시트 데이터를 하나로 합쳐줘"
**사용자 표현 변형**:
- "시트들을 통합해줘"
- "모든 데이터를 한 곳에"
- "전체 데이터 정리"

**AI 실행 시퀀스**:
```bash
# 1. 모든 시트 확인
oa excel workbook-info --include-sheets

# 2. 각 시트에서 데이터 읽기
oa excel range-read --sheet "1월" --range "A1" --expand table --output-file "jan_data.csv"
oa excel range-read --sheet "2월" --range "A1" --expand table --output-file "feb_data.csv"
oa excel range-read --sheet "3월" --range "A1" --expand table --output-file "mar_data.csv"

# 3. 통합 시트 생성
oa excel sheet-add --name "통합데이터"

# 4. 순차적으로 데이터 추가 (헤더는 한 번만)
oa excel table-write --sheet "통합데이터" --range "A1" --data-file "jan_data.csv"
oa excel table-write --sheet "통합데이터" --range "A13" --data-file "feb_data.csv" --skip-header
oa excel table-write --sheet "통합데이터" --range "A25" --data-file "mar_data.csv" --skip-header
```

#### 📈 "피벗테이블로 요약해줘"
**사용자 표현 변형**:
- "요약표 만들어줘"
- "카테고리별로 정리"
- "집계해줘"

**AI 실행 시퀀스**:
```bash
# 1. 피벗테이블 생성
oa excel pivot-create --source-range "A1" --expand table --dest-sheet "요약" --dest-range "A1"

# 2. 피벗테이블 설정
oa excel pivot-configure --pivot-name "PivotTable1" \
  --row-fields "카테고리,제품명" \
  --value-fields "매출액:Sum,수량:Count" \
  --column-fields "분기"

# 3. 피벗차트 생성
oa excel chart-pivot-create --pivot-name "PivotTable1" --chart-type "column" --title "카테고리별 매출 요약"
```

### 보고서 서식 관련

#### 📊 "차트 추가해서 보고서 만들어줘"
**사용자 표현 변형**:
- "그래프가 있는 보고서"
- "시각적인 보고서"
- "차트와 표가 함께"

**AI 실행 시퀀스**:
```bash
# 1. 보고서 시트 생성
oa excel sheet-add --name "월간보고서"

# 2. 제목 및 기본 구조 생성
oa excel range-write --sheet "월간보고서" --range "A1" --data "월간 판매 보고서"

# 3. 요약 데이터 테이블
oa excel table-write --sheet "월간보고서" --range "A3" --data-file "summary_data.csv"

# 4. 차트 추가 (자동 배치)
oa excel chart-add --sheet "월간보고서" --data-range "A3:C15" --chart-type "column" --title "월별 매출 현황" --auto-position

# 5. 추가 차트
oa excel chart-add --sheet "월간보고서" --data-range "E3:F10" --chart-type "pie" --title "제품별 점유율" --auto-position --preferred-position "bottom"
```

## 🔄 자동화 요청

### 반복 작업 관련

#### 🔁 "매달 반복하는 작업 자동화하고 싶어"
**사용자 표현 변형**:
- "매번 하는 일을 자동으로"
- "반복 업무 해결"
- "루틴 작업 자동화"

**AI 대화 패턴**:
```
AI: "매달 어떤 작업을 반복하시나요? 예를 들어:
- 데이터 수집 및 정리
- 보고서 생성
- 차트 업데이트
- 이메일 발송용 자료 만들기

구체적으로 설명해주시면 자동화 방안을 제안해드릴게요."

사용자 답변에 따른 맞춤형 자동화 구현
```

#### 📁 "여러 파일에 동일한 작업 적용해줘"
**사용자 표현 변형**:
- "파일들에 일괄 적용"
- "모든 파일에 같은 처리"
- "배치 작업"

**AI 실행 패턴**:
```bash
# 파일별 순차 처리 예시
for file in ["file1.xlsx", "file2.xlsx", "file3.xlsx"]:
    oa excel workbook-open --file-path "$file"
    oa excel range-convert --range "A1:Z100" --remove-comma
    oa excel chart-add --data-range "A1:C10" --chart-type "line" --auto-position
    # 다음 파일로 이동
```

### 템플릿 관련

#### 📋 "템플릿 만들어서 재사용하고 싶어"
**사용자 표현 변형**:
- "양식 만들어줘"
- "틀을 만들어서"
- "표준 서식"

**AI 실행 시퀀스**:
```bash
# 1. 템플릿 워크북 생성
oa excel workbook-create --name "월간보고서_템플릿" --save-path "templates/monthly_template.xlsx"

# 2. 기본 구조 설정
oa excel range-write --range "A1" --data "월간 보고서"
oa excel range-write --range "A3:D3" --data '["날짜", "제품", "매출", "비고"]'

# 3. 표 서식 적용
oa excel table-create --range "A3:D20" --table-name "데이터입력" --table-style "TableStyleMedium2"

# 4. 차트 영역 설정
oa excel chart-add --data-range "A3:C10" --chart-type "column" --title "월별 실적" --position "F3"

# 5. 안내 메시지
oa excel range-write --range "A25" --data "이 템플릿에 데이터를 입력하고 oa excel pivot-refresh를 실행하세요"
```

## 🔧 기술적 문제 해결

### 파일 및 연결 문제

#### 🔗 "파일이 안 열려"
**사용자 표현 변형**:
- "에러가 나"
- "접근이 안 돼"
- "파일을 찾을 수 없다고 해"

**AI 대응 시퀀스**:
```bash
# 1. 현재 상황 파악
oa excel workbook-list --detailed

# 2. 파일 경로 확인 및 대안 제시
# 절대 경로로 다시 시도
oa excel workbook-open --file-path "C:\Users\사용자명\Documents\파일명.xlsx"

# 3. 새 파일 생성 옵션 제시
oa excel workbook-create --name "새작업파일" --save-path "new_work.xlsx"
```

#### ⚠️ "명령어를 모르겠어"
**AI 대응 패턴**:
```
AI: "어떤 작업을 하고 싶으신지 자연스럽게 말씀해주세요.
예를 들어 '차트 만들어줘', '데이터 정리해줘' 같이요.
제가 알아서 적절한 명령어를 선택해서 실행해드릴게요."

# 그 후 사용자 의도에 맞는 명령어 자동 선택
```

## 💡 고급 활용 요청

### 대시보드 관련

#### 📊 "대시보드 만들어줘"
**사용자 표현 변형**:
- "종합 현황판"
- "한눈에 보는 차트들"
- "경영진 보고용 화면"

**AI 실행 시퀀스**:
```bash
# 1. 대시보드 시트 생성
oa excel sheet-add --name "Dashboard"

# 2. 여러 데이터 소스에서 요약 데이터 생성
oa excel pivot-create --source-range "RawData!A1" --expand table --dest-sheet "Dashboard" --dest-range "A20"
oa excel pivot-configure --pivot-name "PivotTable1" --row-fields "카테고리" --value-fields "매출:Sum"

# 3. 여러 차트 자동 배치
oa excel chart-add --sheet "Dashboard" --data-range "B2:C8" --chart-type "column" --title "카테고리별 매출" --position "B2"
oa excel chart-add --sheet "Dashboard" --data-range "E2:F6" --chart-type "pie" --title "지역별 분포" --auto-position
oa excel chart-add --sheet "Dashboard" --data-range "A20:B25" --chart-type "line" --title "월별 추세" --auto-position --preferred-position "bottom"

# 4. 슬라이서 추가 (인터랙티브 필터)
oa excel slicer-add --pivot-name "PivotTable1" --field "날짜" --position "F20"
```

### 복합 분석

#### 📈 "상관관계 분석해줘"
**사용자 표현 변형**:
- "관련성 봐줘"
- "연관성 분석"
- "어떤 요인이 영향을 주는지"

**AI 실행 시퀀스**:
```bash
# 1. 산점도 차트 생성
oa excel chart-add --data-range "A1:B100" --chart-type "scatter" --title "가격 vs 판매량 상관관계"

# 2. 추세선 추가 (Excel 기능)
oa excel chart-configure --chart-name "Chart1" --add-trendline --trendline-type "linear"

# 3. 상관계수 계산
oa excel range-write --range "E1:F3" --data '[
  ["상관계수", "=CORREL(A:A, B:B)"],
  ["R제곱", "=RSQ(A:A, B:B)"],
  ["표준편차", "=STDEV(B:B)"]
]'
```

---

## 🎯 AI 에이전트 사용 팁

### 1. 패턴 인식
- 사용자가 "차트", "그래프" 언급 → 데이터 위치부터 확인
- "정리", "정리" 언급 → 어떤 종류의 정리인지 구체화
- "자동화" 언급 → 현재 수동 프로세스부터 파악

### 2. 단계별 실행
- 복잡한 요청은 여러 단계로 분할
- 각 단계마다 결과 확인 및 설명
- 다음 단계 진행 전 사용자 확인

### 3. 백업 및 안전성
- 원본 데이터 보존 (새 시트 활용)
- 실행 전 현재 상황 설명
- 실패 시 대안 방법 제시

### 4. 사용자 경험 개선
- 전문 용어 → 일반 용어로 번역
- 실행 결과를 이해하기 쉽게 해석
- 추가로 할 수 있는 작업 제안

이 예시들을 참고하여 사용자의 다양한 요청에 효과적으로 대응하세요.