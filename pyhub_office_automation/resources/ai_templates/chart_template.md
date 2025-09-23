## 차트 제안 예시

### 차트 선택 가이드

**`chart-add` 사용 권장 상황:**
- 간단한 데이터 시각화
- 크로스 플랫폼 호환성 필요
- 빠른 차트 생성
- 피벗차트 타임아웃 문제 회피

**`chart-pivot-create` 사용 상황 (Windows 전용):**
- 대화형 필터링 기능 필요
- 복잡한 데이터 집계
- `--skip-pivot-link` 옵션 사용 권장

### 차트 유형별 예시

#### 1. 판매량 비교 (막대형 차트)
```bash
oa excel chart-add \
  --sheet "데이터" \
  --data-range "A1:B10" \
  --chart-type "Column" \
  --title "제품별 판매량" \
  --x-axis-title "제품명" \
  --y-axis-title "판매량(개)"
```

**권장 용도**: 카테고리별 수치 비교
- 제품별 판매량
- 지역별 매출
- 월별 실적 비교

#### 2. 시간 추세 (선형 차트)
```bash
oa excel chart-add \
  --sheet "데이터" \
  --data-range "A1:B20" \
  --chart-type "Line" \
  --title "월별 매출 추이" \
  --x-axis-title "월" \
  --y-axis-title "매출(만원)"
```

**권장 용도**: 시간에 따른 변화 추적
- 월별/일별 추이
- 성장률 분석
- 계절성 패턴

#### 3. 구성 비율 (원형 차트)
```bash
oa excel chart-add \
  --sheet "데이터" \
  --data-range "A1:B6" \
  --chart-type "Pie" \
  --title "시장 점유율" \
  --show-data-labels
```

**권장 용도**: 전체 대비 비율 표시
- 시장 점유율
- 예산 구성
- 고객 분포

### 피벗테이블 기반 차트

#### 피벗테이블 구성 요소
- **행 영역**: 카테고리 분류 (제품, 지역, 날짜 등)
- **열 영역**: 추가 분류 축 (연도, 분기 등)
- **값 영역**: 집계할 수치 (매출, 수량, 평균 등)
- **필터 영역**: 데이터 필터링 조건

#### 피벗차트 생성 예시
```bash
oa excel chart-pivot-create \
  --sheet "원본데이터" \
  --data-range "A1:E1000" \
  --rows "지역,제품" \
  --values "매출액:합계" \
  --chart-type "Column" \
  --skip-pivot-link \
  --pivot-table-name "Sales_Analysis"
```

### 차트 커스터마이징

```bash
# 차트 설정 변경
oa excel chart-configure \
  --chart-name "Chart1" \
  --title "새 제목" \
  --show-legend \
  --legend-position "Right"

# 차트 위치 조정
oa excel chart-position \
  --chart-name "Chart1" \
  --left 100 \
  --top 50 \
  --width 400 \
  --height 300

# 차트 내보내기
oa excel chart-export \
  --chart-name "Chart1" \
  --output-path "chart.png" \
  --format "PNG"
```

### 차트 제안 템플릿

1. **게임별 글로벌 판매량 (막대형)**: 각 게임의 글로벌 판매량(백만장)을 내림차순으로 하고, 한 눈에 베스트셀러 규모 차이를 파악
   - **인사이트**: 상위 3개 게임이 전체 매출의 60% 차지
   - **피벗테이블 구성**: 게임명(행), 판매량 합계(값), 내림차순 정렬
   - **차트 설정**: Column 차트, 제목 "글로벌 게임 판매량 TOP 10"

2. **지역별 월별 매출 추이 (선형)**: 각 지역의 월별 매출 변화를 추적하여 계절성 패턴 분석
   - **인사이트**: 12월 매출 급증, 2월 매출 저조
   - **피벗테이블 구성**: 월(행), 지역(열), 매출액 합계(값)
   - **차트 설정**: Line 차트, 범례 표시, 격자선 활성화

3. **제품 카테고리별 이익률 (원형)**: 전체 이익에서 각 카테고리가 차지하는 비중 시각화
   - **인사이트**: 모바일 게임이 이익의 45% 차지
   - **피벗테이블 구성**: 카테고리(행), 이익률 평균(값)
   - **차트 설정**: Pie 차트, 데이터 레이블 표시, 퍼센트 형식