# 고급 기능 가이드

## 목차
- [Map Chart 활용 가이드 (Issue #72)](#map-chart-활용-가이드-issue-72)
  - [지리 데이터 시각화 전략](#지리-데이터-시각화-전략)
  - [실전 활용 시나리오](#실전-활용-시나리오)
- [차트 제안 가이드](#차트-제안-가이드)
  - [차트 선택 가이드](#차트-선택-가이드)
  - [차트 유형별 예시](#차트-유형별-예시)
  - [피벗테이블 기반 차트](#피벗테이블-기반-차트)
  - [차트 커스터마이징](#차트-커스터마이징)

---

## Map Chart 활용 가이드 (Issue #72)

### 지리 데이터 시각화 전략

Claude Code의 체계적 접근을 활용한 Map Chart 워크플로우:

#### 1단계: 데이터 준비 및 위치명 검증

```bash
# 위치명 형식 확인 (25개 구 가이드)
oa excel map-location-guide --region seoul --show-all

# 데이터 내 위치명 자동 변환 테스트
oa excel map-location-guide --test "강남구,서초구,songpa" --format json

# Python 시각화용 데이터 검증
oa excel map-visualize --data-file district_sales.csv --validate-only
```

**Claude의 데이터 품질 검증 포인트:**
- ✅ 위치명 일관성 확인 (한글/영문 혼용 여부)
- ✅ 25개 구 전체 커버리지 검증
- ✅ 결측값 및 이상치 탐지
- ✅ Excel vs Python 시각화 적합성 판단

#### 2단계: 시각화 방법 선택

**의사결정 트리:**

```python
def select_visualization_method(requirements):
    """Claude가 추천하는 시각화 방법 선택 로직"""

    # Excel 환경 확인
    if requirements.excel_available and requirements.microsoft_365:
        if requirements.interactive_filters:
            return "Excel Map Chart"  # Excel 대시보드용
        elif requirements.simple_visual:
            return "Excel Map Chart"  # 간단한 보고서용

    # Python 환경 추천
    if requirements.cross_platform or not requirements.excel_available:
        if requirements.offline_html:
            return "Python folium (choropleth)"  # 대화형 HTML
        elif requirements.point_markers:
            return "Python folium (marker)"  # 핀 마커 지도

    # 기본 추천
    return "Python folium"  # 가장 범용적
```

#### 3단계: Excel Map Chart 구현 (Phase 1)

```bash
# 데이터가 Excel Table 형식일 때
oa excel shell --workbook-name "sales_data.xlsx"

[Excel: sales_data.xlsx > Sheet1] > table-list  # 테이블 구조 확인
[Excel: sales_data.xlsx > Sheet1] > use sheet "DistrictSales"

# 위치명 자동 변환 확인
[Excel: sales_data.xlsx > DistrictSales] > range-read --range "A1:A26"
# → 강남구, 서초구 등 한글 확인

# Map Chart 생성
[Excel: sales_data.xlsx > DistrictSales] > chart-add \
  --data-range "A1:B26" \
  --chart-type "map" \
  --title "서울시 구별 매출 분포" \
  --auto-position

[Excel: sales_data.xlsx > DistrictSales] > exit
```

**Excel Map Chart 장점:**
- ✅ Excel 환경 내 완전 통합
- ✅ PowerPivot 연동 가능
- ✅ Bing Maps 실시간 업데이트

**제약사항:**
- ❌ Windows + Excel 2016+ 필수
- ❌ 인터넷 연결 필요
- ❌ 커스텀 지도 제한적

#### 4단계: Python folium 구현 (Phase 3 - 권장)

```bash
# 데이터 추출 (Excel → CSV)
oa excel table-read --table-name "DistrictSales" --output-file sales_by_district.csv

# Choropleth 지도 생성
oa excel map-visualize \
  --data-file sales_by_district.csv \
  --value-column "sales_amount" \
  --title "서울시 구별 매출 분포" \
  --color-scheme YlOrRd \
  --output-file seoul_sales_map.html

# 브라우저에서 확인 (자동 위치명 변환됨)
# "강남구" → "Seoul Gangnam" (자동)
```

**Python 시각화 장점:**
- ✅ Excel 불필요 - CI/CD 파이프라인 통합 가능
- ✅ 크로스 플랫폼 - Linux 서버에서도 실행
- ✅ 버전 관리 - HTML 파일로 결과물 저장
- ✅ 자동화 친화적 - 배치 스크립트로 정기 업데이트

#### 5단계: 고급 워크플로우 - 다중 데이터셋 비교

```bash
# 시나리오: Q1, Q2, Q3, Q4 분기별 지도 자동 생성

# Shell Mode 활용
oa excel shell

[Excel: None > None] > use workbook "quarterly_data.xlsx"
[Excel: quarterly_data.xlsx > Sheet1] > use sheet Q1
[Excel: quarterly_data.xlsx > Q1] > table-read --output-file q1.csv

# Python 지도 생성
[Excel: quarterly_data.xlsx > Q1] > !oa excel map-visualize \
  --data-file q1.csv \
  --title "Q1 2024 서울시 구별 매출" \
  --output-file q1_map.html

# Q2, Q3, Q4 반복
[Excel: quarterly_data.xlsx > Q1] > use sheet Q2
[Excel: quarterly_data.xlsx > Q2] > table-read --output-file q2.csv
# ... (반복)
```

#### Claude의 Map Chart 품질 검증 체크리스트

**데이터 무결성:**
- [ ] 25개 구 전체 데이터 존재 확인
- [ ] 위치명 일관성 (혼용 방지)
- [ ] 결측값 0개 또는 명시적 처리
- [ ] 값 범위 이상치 검증 (예: 음수 매출)

**시각화 품질:**
- [ ] 색상 스킴이 데이터 의미와 일치 (빨강=높음, 파랑=낮음)
- [ ] 범례 및 툴팁 가독성
- [ ] 지도 중심 및 줌 레벨 최적화
- [ ] 모바일 반응형 (HTML 출력 시)

**사용자 경험:**
- [ ] 로딩 시간 1초 이내 (Python HTML)
- [ ] 상호작용 즉시 반응 (클릭, 호버)
- [ ] 접근성 (색각이상 고려)
- [ ] 다국어 레이블 (한글/영문 병기)

### Map Chart 에러 처리 패턴

```python
# Claude 권장 에러 처리 워크플로우

def robust_map_visualization():
    """안전한 지도 시각화 파이프라인"""

    try:
        # 1단계: 데이터 검증
        result = subprocess.run(
            ["oa", "excel", "map-visualize",
             "--data-file", "sales.csv",
             "--validate-only", "--format", "json"],
            capture_output=True, text=True, check=True
        )

        validation = json.loads(result.stdout)

        if validation["data"]["unmatched_count"] > 0:
            print(f"⚠️  {validation['data']['unmatched_count']} locations need fixing")
            # Claude가 자동으로 위치명 수정 제안
            for item in validation["data"]["unmatched"]:
                print(f"  - {item['input']}: {item['suggestions'][0]}")
            return False

        # 2단계: 지도 생성
        result = subprocess.run(
            ["oa", "excel", "map-visualize",
             "--data-file", "sales.csv",
             "--value-column", "amount",
             "--output-file", "output.html"],
            capture_output=True, text=True, check=True
        )

        print("✓ Map created successfully")
        return True

    except subprocess.CalledProcessError as e:
        print(f"Error: {e.stderr}")
        # Claude가 에러 원인 분석 및 해결책 제시
        return False
```

### 실전 활용 시나리오

**1. 부동산 시장 분석 대시보드**
```bash
# 평균 매매가, 전세가, 월세 3개 지도 동시 생성
# 결과:
# - seoul_sales_price_map.html
# - seoul_jeonse_price_map.html
# - seoul_monthly_rent_map.html
```

**2. 인구 통계 시계열 분석**
```bash
# 2020-2024년 연도별 인구 변화 애니메이션
# (Claude가 연도별 HTML 생성 후 슬라이드쇼 스크립트 제안)
```

**3. 공공데이터 시각화 자동화**
```bash
# 서울시 열린데이터광장 API → CSV → 지도 자동 업데이트
# cron: 매일 오전 9시 자동 실행
```

---

## 차트 제안 가이드

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

**차트 선택 의사결정 로직:**
1. **Data Size**: Large datasets (>1000 rows) → `chart-add` (due to timeout issues)
2. **Interactivity**: Need filtering/drilling → Use pivot table + `chart-add` separately
3. **Platform**: macOS environment → `chart-add` only
4. **Complexity**: Simple visualization → `chart-add`
5. **Existing Pivot**: Pivot table already exists → `chart-add` with pivot data range

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
  --name "Chart1" \
  --title "새 제목" \
  --show-legend \
  --legend-position "Right"

# 차트 위치 조정
oa excel chart-position \
  --name "Chart1" \
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

#### 1. 게임별 글로벌 판매량 (막대형)
- **목적**: 각 게임의 글로벌 판매량(백만장)을 내림차순으로 하고, 한 눈에 베스트셀러 규모 차이를 파악
- **인사이트**: 상위 3개 게임이 전체 매출의 60% 차지
- **피벗테이블 구성**: 게임명(행), 판매량 합계(값), 내림차순 정렬
- **차트 설정**: Column 차트, 제목 "글로벌 게임 판매량 TOP 10"

```bash
oa excel chart-add \
  --sheet "데이터" \
  --data-range "GameData[글로벌 판매량]" \
  --chart-type "Column" \
  --title "글로벌 게임 판매량 TOP 10"
```

#### 2. 지역별 월별 매출 추이 (선형)
- **목적**: 각 지역의 월별 매출 변화를 추적하여 계절성 패턴 분석
- **인사이트**: 12월 매출 급증, 2월 매출 저조
- **피벗테이블 구성**: 월(행), 지역(열), 매출액 합계(값)
- **차트 설정**: Line 차트, 범례 표시, 격자선 활성화

```bash
oa excel chart-add \
  --sheet "데이터" \
  --data-range "A1:E13" \
  --chart-type "Line" \
  --title "지역별 월별 매출 추이" \
  --show-legend
```

#### 3. 제품 카테고리별 이익률 (원형)
- **목적**: 전체 이익에서 각 카테고리가 차지하는 비중 시각화
- **인사이트**: 모바일 게임이 이익의 45% 차지
- **피벗테이블 구성**: 카테고리(행), 이익률 평균(값)
- **차트 설정**: Pie 차트, 데이터 레이블 표시, 퍼센트 형식

```bash
oa excel chart-add \
  --sheet "데이터" \
  --data-range "A1:B6" \
  --chart-type "Pie" \
  --title "제품 카테고리별 이익률" \
  --show-data-labels
```

---

## 참고 문서

- [CLAUDE.md](../CLAUDE.md) - AI Agent Quick Reference
- [SHELL_USER_GUIDE.md](./SHELL_USER_GUIDE.md) - Shell Mode 완벽 가이드
- [CLAUDE_CODE_PATTERNS.md](./CLAUDE_CODE_PATTERNS.md) - Claude Code 특화 패턴
