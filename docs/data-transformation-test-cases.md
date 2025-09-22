# 데이터 변환 테스트 케이스

## 개요

`data-analyze`와 `data-transform` 명령어의 각 데이터 형식 문제별 테스트 케이스를 정의합니다.

## 테스트 케이스 구조

각 테스트 케이스는 다음 구조를 따릅니다:
- **입력 데이터**: 문제가 있는 원본 데이터
- **예상 분석 결과**: data-analyze 명령어 출력
- **변환 방법**: 적용할 변환 타입
- **예상 변환 결과**: data-transform 명령어 출력
- **검증 방법**: 결과 검증 로직

## 테스트 케이스 1: Cross-tab 형식 (교차표)

### 입력 데이터
```
제품명     | 1월  | 2월  | 3월  | 4월  | 5월  | 6월
---------|-----|-----|-----|-----|-----|-----
노트북    | 100 | 120 | 150 | 130 | 140 | 160
마우스    | 50  | 45  | 60  | 55  | 70  | 65
키보드    | 30  | 35  | 40  | 32  | 38  | 42
```

### 예상 분석 결과
```json
{
  "format_type": "cross_tab",
  "pivot_ready": false,
  "issues": [{
    "type": "cross_tab_format",
    "severity": "high",
    "description": "월별 데이터가 열로 배치되어 피벗테이블에 부적합",
    "affected_range": "B1:G1",
    "solution": "unpivot"
  }],
  "transformation_impact": {
    "current_shape": [4, 7],
    "expected_shape": [18, 3]
  }
}
```

### 변환 명령어
```bash
oa excel data-transform --source-range "A1:G4" \
  --transform-type "unpivot" \
  --id-columns "제품명" \
  --output-sheet "Transformed"
```

### 예상 변환 결과
```
제품명 | 월  | 매출액
------|----|-----
노트북 | 1월 | 100
노트북 | 2월 | 120
노트북 | 3월 | 150
...    | ... | ...
```

### 검증 코드
```python
def test_cross_tab_transformation():
    # 원본 데이터 생성
    original_data = create_cross_tab_data()

    # 분석 실행
    analysis = run_data_analyze(original_data)
    assert analysis['format_type'] == 'cross_tab'
    assert not analysis['pivot_ready']

    # 변환 실행
    transformed = run_data_transform(original_data, 'unpivot')
    assert transformed.shape == (18, 3)
    assert list(transformed.columns) == ['제품명', '월', '매출액']

    # 피벗테이블 적합성 검증
    pivot_test = run_data_analyze(transformed)
    assert pivot_test['pivot_ready'] == True
```

## 테스트 케이스 2: Multi-level Headers (다단계 헤더)

### 입력 데이터
```
Row 1: |      | 2024년        |        | 2025년        |        |
Row 2: |      | Q1    | Q2    | Q3    | Q1    | Q2    |
Row 3: | 제품 | 1월   | 2월   | 3월   | 1월   | 2월   |
Row 4: | A    | 100   | 120   | 150   | 110   | 130   |
```

### 예상 분석 결과
```json
{
  "format_type": "multi_level_headers",
  "header_rows": 3,
  "issues": [{
    "type": "multi_level_headers",
    "severity": "high",
    "description": "3단계 헤더 구조로 인해 필드 인식 불가",
    "affected_range": "A1:F3",
    "solution": "flatten-headers"
  }]
}
```

### 변환 명령어
```bash
oa excel data-transform --source-range "A1:F5" \
  --transform-type "flatten-headers" \
  --output-sheet "Flattened"
```

### 예상 변환 결과
```
제품 | 2024년_Q1_1월 | 2024년_Q1_2월 | 2024년_Q2_3월 | 2025년_Q1_1월 | 2025년_Q1_2월
----|-------------|-------------|-------------|-------------|-------------
A   | 100         | 120         | 150         | 110         | 130
```

## 테스트 케이스 3: Merged Cells (병합된 셀)

### 입력 데이터
```
지역    | 제품   | 매출액
------- |-------|-------
서울    | 노트북  | 1000
(병합)   | 마우스  | 500
(병합)   | 키보드  | 300
부산    | 노트북  | 800
(병합)   | 마우스  | 400
```

### 예상 분석 결과
```json
{
  "issues": [{
    "type": "merged_cells",
    "severity": "medium",
    "description": "5개의 병합된 셀로 인해 데이터 불완전",
    "affected_range": "A2:A6",
    "solution": "unmerge"
  }]
}
```

### 변환 명령어
```bash
oa excel data-transform --source-range "A1:C6" \
  --transform-type "unmerge" \
  --output-sheet "Unmerged"
```

### 예상 변환 결과
```
지역 | 제품   | 매출액
----|-------|-------
서울 | 노트북  | 1000
서울 | 마우스  | 500
서울 | 키보드  | 300
부산 | 노트북  | 800
부산 | 마우스  | 400
```

## 테스트 케이스 4: Subtotals Mixed (소계 혼재)

### 입력 데이터
```
지역 | 제품   | 매출액
----|-------|-------
서울 | 노트북  | 1000
서울 | 마우스  | 500
서울 | 소계   | 1500
부산 | 노트북  | 800
부산 | 마우스  | 400
부산 | 소계   | 1200
총계 |       | 2700
```

### 예상 분석 결과
```json
{
  "issues": [{
    "type": "subtotals_present",
    "severity": "medium",
    "description": "3개의 소계 행이 데이터와 혼재",
    "affected_range": "A3,A6,A7",
    "solution": "remove-subtotals"
  }]
}
```

### 변환 명령어
```bash
oa excel data-transform --source-range "A1:C8" \
  --transform-type "remove-subtotals" \
  --output-sheet "NoSubtotals"
```

### 예상 변환 결과
```
지역 | 제품   | 매출액
----|-------|-------
서울 | 노트북  | 1000
서울 | 마우스  | 500
부산 | 노트북  | 800
부산 | 마우스  | 400
```

## 테스트 케이스 5: Wide Format (넓은 형식)

### 입력 데이터
```
제품   | 1월_매출 | 1월_비용 | 2월_매출 | 2월_비용 | 3월_매출 | 3월_비용
------|--------|--------|--------|--------|--------|--------
노트북 | 1000   | 800    | 1200   | 900    | 1100   | 850
마우스 | 500    | 300    | 550    | 320    | 600    | 350
```

### 예상 분석 결과
```json
{
  "format_type": "wide_format",
  "issues": [{
    "type": "wide_format",
    "severity": "high",
    "description": "여러 지표가 열로 분산되어 분석 어려움",
    "solution": "unpivot"
  }]
}
```

### 변환 명령어
```bash
oa excel data-transform --source-range "A1:G3" \
  --transform-type "unpivot" \
  --id-columns "제품" \
  --output-sheet "LongFormat"
```

### 예상 변환 결과
```
제품   | 지표     | 월  | 값
------|---------|-------|----
노트북 | 매출     | 1월   | 1000
노트북 | 비용     | 1월   | 800
노트북 | 매출     | 2월   | 1200
...   | ...     | ...   | ...
```

## 테스트 케이스 6: Auto Transform (복합 문제)

### 입력 데이터
```
Row 1: | 지역    |        | 2024년 매출      |        |
Row 2: | (병합)   | 제품    | 1월    | 2월    | 3월    |
Row 3: | 서울    | 노트북   | 1000   | 1200   | 1100   |
Row 4: | (병합)   | 마우스   | 500    | 550    | 600    |
Row 5: | (병합)   | 소계    | 1500   | 1750   | 1700   |
```

### 예상 분석 결과
```json
{
  "format_type": "complex_multi_issue",
  "issues": [
    {"type": "merged_cells", "solution": "unmerge"},
    {"type": "multi_level_headers", "solution": "flatten-headers"},
    {"type": "cross_tab_format", "solution": "unpivot"},
    {"type": "subtotals_present", "solution": "remove-subtotals"}
  ]
}
```

### 변환 명령어
```bash
oa excel data-transform --source-range "A1:E5" \
  --transform-type "auto" \
  --output-sheet "AutoTransformed"
```

### 예상 변환 결과
```
지역 | 제품   | 월  | 매출액
----|-------|-------|-------
서울 | 노트북 | 1월   | 1000
서울 | 노트북 | 2월   | 1200
서울 | 노트북 | 3월   | 1100
서울 | 마우스 | 1월   | 500
서울 | 마우스 | 2월   | 550
서울 | 마우스 | 3월   | 600
```

## AI 통합 테스트 케이스

### 테스트 케이스 7: AI 보조 분석

```bash
# Codex를 활용한 복잡한 분석
oa excel data-analyze --file-path "complex.xlsx" --ai-assist
```

### 예상 AI 분석 결과
```json
{
  "ai_analysis": {
    "strategy": "1단계: 병합 해제, 2단계: 헤더 평탄화, 3단계: unpivot",
    "code_suggestion": "pd.melt() 사용 후 fillna() 적용",
    "complexity_score": 0.8,
    "estimated_processing_time": "10-15초"
  }
}
```

## 성능 테스트 케이스

### 대용량 데이터 테스트
- **크기**: 10,000행 × 50열
- **목표 성능**: 분석 5초 이내, 변환 15초 이내
- **메모리 사용량**: 500MB 이하

### 스트레스 테스트
- **동시 실행**: 5개 명령어 병렬 실행
- **다양한 형식**: 각 문제 타입별 동시 처리
- **리소스 모니터링**: CPU, 메모리 사용률 추적

## 테스트 자동화

### pytest 구조
```python
# tests/test_data_commands.py

class TestDataAnalyze:
    def test_cross_tab_detection(self):
        # 교차표 형식 감지 테스트
        pass

    def test_merged_cells_detection(self):
        # 병합된 셀 감지 테스트
        pass

class TestDataTransform:
    def test_unpivot_transformation(self):
        # unpivot 변환 테스트
        pass

    def test_auto_transformation(self):
        # 자동 변환 테스트
        pass

class TestAIIntegration:
    def test_codex_analysis(self):
        # Codex 통합 테스트
        pass

    def test_ai_fallback(self):
        # AI 도구 실패 시 폴백 테스트
        pass
```

### 테스트 데이터 생성기
```python
def generate_test_data(data_type, size=(100, 10)):
    """테스트용 문제 데이터 생성"""
    if data_type == 'cross_tab':
        return create_cross_tab_sample(size)
    elif data_type == 'merged_cells':
        return create_merged_cells_sample(size)
    # ... 기타 타입들
```

이러한 종합적인 테스트 케이스를 통해 `data-analyze`와 `data-transform` 명령어의 정확성과 신뢰성을 보장할 수 있습니다.