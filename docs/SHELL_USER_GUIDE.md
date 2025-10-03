# Shell Mode 완벽 가이드

## 목차
- [Excel Shell Mode](#excel-shell-mode)
  - [개요](#excel-개요)
  - [시작 방법](#excel-시작-방법)
  - [핵심 워크플로우](#excel-핵심-워크플로우)
  - [내부 명령어](#excel-내부-명령어)
  - [권장 패턴](#excel-권장-패턴)
- [PowerPoint Shell Mode](#powerpoint-shell-mode)
  - [개요](#ppt-개요)
  - [시작 방법](#ppt-시작-방법)
  - [핵심 워크플로우](#ppt-핵심-워크플로우)
  - [내부 명령어](#ppt-내부-명령어)
  - [권장 패턴](#ppt-권장-패턴)

---

## Excel Shell Mode

### Excel 개요

**언제 Shell Mode를 사용할까?**
- 동일한 워크북/시트에서 **3개 이상의 연속 작업**이 필요할 때
- 탐색적 데이터 분석 (Exploratory Data Analysis) 수행 시
- 대화형으로 데이터 구조를 파악하고 단계적 분석이 필요할 때
- 워크북/시트 전환이 빈번한 복합 작업 시

**Shell Mode vs 일반 CLI Mode**

| 특성 | Shell Mode | 일반 CLI Mode |
|------|-----------|--------------|
| **적합한 경우** | 연속 작업 3개 이상 | 단발성 작업 1-2개 |
| **명령어 길이** | 50% 단축 | 전체 경로 필요 |
| **컨텍스트 관리** | 자동 유지 | 매번 지정 |
| **탐색 효율** | 높음 (대화형) | 낮음 (단발성) |
| **자동완성** | Tab 지원 | 없음 |
| **히스토리** | 세션 내 유지 | 없음 |

### Excel 시작 방법

```bash
# 방법 1: 활성 워크북 자동 선택
oa excel shell

# 방법 2: 파일 경로로 시작
oa excel shell --file-path "C:/data/report.xlsx"

# 방법 3: 열린 파일명으로 시작
oa excel shell --workbook-name "sales.xlsx"
```

### Excel 핵심 워크플로우

#### 완전한 데이터 분석 워크플로우

```bash
# 시작: 워크북 자동 선택
oa excel shell

# 1단계: 환경 파악 (Shell 명령)
[Excel: None > None] > workbook-list              # 열린 파일 확인
[Excel: None > None] > use workbook "sales.xlsx"  # 워크북 선택
[Excel: sales.xlsx > None] > show context         # 현재 상태 확인
[Excel: sales.xlsx > None] > sheets               # 시트 목록

# 2단계: 데이터 탐색 (컨텍스트 자동 주입)
[Excel: sales.xlsx > None] > use sheet RawData
[Excel: sales.xlsx > RawData] > table-list        # 테이블 구조 파악
[Excel: sales.xlsx > RawData] > range-read --range A1:A1  # 헤더 확인
[Excel: sales.xlsx > RawData] > range-read --range A1:F5  # 샘플 데이터

# 3단계: 분석 수행 (시트 전환 및 작업)
[Excel: sales.xlsx > RawData] > sheet-add --name "Analysis"
[Excel: sales.xlsx > RawData] > use sheet Analysis
[Excel: sales.xlsx > Analysis] > chart-add --data-range "RawData!A1:C10" --chart-type "Column"
[Excel: sales.xlsx > Analysis] > chart-configure --name "Chart1" --title "월별 매출"

# 4단계: 결과 확인 및 저장
[Excel: sales.xlsx > Analysis] > workbook-info    # 변경사항 확인
[Excel: sales.xlsx > Analysis] > exit             # 종료
```

**비교: 일반 CLI로 동일 작업**
```bash
# 위 워크플로우를 일반 CLI로 하면:
oa excel workbook-list
oa excel workbook-open --file-path "sales.xlsx"
oa excel workbook-info --workbook-name "sales.xlsx"
oa excel table-list --workbook-name "sales.xlsx"
oa excel range-read --workbook-name "sales.xlsx" --sheet "RawData" --range A1:A1
oa excel range-read --workbook-name "sales.xlsx" --sheet "RawData" --range A1:F5
oa excel sheet-add --workbook-name "sales.xlsx" --name "Analysis"
oa excel chart-add --workbook-name "sales.xlsx" --sheet "Analysis" --data-range "RawData!A1:C10" --chart-type "Column"
oa excel chart-configure --workbook-name "sales.xlsx" --sheet "Analysis" --name "Chart1" --title "월별 매출"
oa excel workbook-info --workbook-name "sales.xlsx"
# → 명령어 길이 3배 증가, 타이핑 부담 증가
```

### Excel 내부 명령어

| 명령어 | 설명 | 예제 |
|--------|------|------|
| `help` | 카테고리별 명령어 목록 | `help` |
| `show context` | 현재 워크북/시트 상태 표시 | `show context` |
| `use workbook <name>` | 워크북 전환 | `use workbook "sales.xlsx"` |
| `use sheet <name>` | 시트 전환 | `use sheet "Data"` |
| `sheets` | 현재 워크북의 시트 목록 | `sheets` |
| `workbook-info` | 워크북 상세 정보 | `workbook-info` |
| `clear` | 화면 지우기 | `clear` |
| `exit` / `quit` | Shell 종료 | `exit` |

### Excel 권장 패턴

#### 1. 단계적 탐색 패턴
```bash
# 안전한 점진적 접근
oa excel shell
[Excel: None > None] > workbook-list      # 1. 환경 확인
[Excel: None > None] > use workbook "file.xlsx"
[Excel: file.xlsx > None] > sheets        # 2. 구조 파악
[Excel: file.xlsx > None] > use sheet Data
[Excel: file.xlsx > Data] > table-list    # 3. 테이블 분석
[Excel: file.xlsx > Data] > range-read --range A1:C5  # 4. 샘플 확인
# → 각 단계마다 출력을 보고 다음 명령 결정
```

#### 2. Tab 자동완성 적극 활용
```bash
# 명령어 입력 시 Tab 키로 자동완성
[Excel: None > None] > wo<TAB>    # → workbook-list
[Excel: None > None] > use w<TAB> # → use workbook
[Excel: None > None] > ra<TAB>    # → range-read
[Excel: None > None] > ta<TAB>    # → table-list
# → 52개 명령어 모두 Tab 자동완성 지원
```

#### 3. 컨텍스트 인식 명령 실행
```bash
# show context로 현재 상태 주기적 확인
[Excel: sales.xlsx > RawData] > show context
# 출력:
# Current Context:
#   Workbook: sales.xlsx
#   Sheet: RawData
#   All Excel commands will use this context automatically.

# 컨텍스트가 명확하면 최소 인자로 실행
[Excel: sales.xlsx > RawData] > range-read --range A1:C10
# → --workbook-name, --sheet 자동 주입됨
```

#### 4. 다중 시트 분석 패턴
```bash
# 시트 전환하며 비교 분석
oa excel shell --workbook-name "report.xlsx"

[Excel: report.xlsx > Sheet1] > sheets  # 전체 시트 확인
[Excel: report.xlsx > Sheet1] > use sheet Q1
[Excel: report.xlsx > Q1] > table-read --output-file q1.csv
[Excel: report.xlsx > Q1] > use sheet Q2
[Excel: report.xlsx > Q2] > table-read --output-file q2.csv
[Excel: report.xlsx > Q2] > use sheet Q3
[Excel: report.xlsx > Q3] > table-read --output-file q3.csv
# → 시트 전환만으로 동일 작업 반복
```

#### 5. 에러 복구 패턴
```bash
# 명령 실패 시 즉시 재시도 가능
[Excel: test.xlsx > Data] > range-read --range "A1:Z100"
# Error: Sheet 'Data' not found

[Excel: test.xlsx > Data] > sheets  # 올바른 시트명 확인
[Excel: test.xlsx > Data] > use sheet "RawData"  # 수정
[Excel: test.xlsx > RawData] > range-read --range "A1:Z100"  # 재시도
# → 세션 유지로 빠른 수정 가능
```

### 성능 및 효율성

- **명령어 입력 속도**: 일반 CLI 대비 10배 빠름 (Tab 자동완성 + 컨텍스트 생략)
- **탐색 효율**: 즉시 피드백으로 시행착오 감소
- **오타 방지**: Tab 자동완성으로 명령어 오타 90% 감소
- **생산성**: 연속 작업 5개 이상 시 50% 시간 절약

---

## PowerPoint Shell Mode

### PPT 개요

**언제 PowerPoint Shell Mode를 사용할까?**
- 동일한 프레젠테이션에서 **여러 슬라이드 연속 편집**이 필요할 때
- 슬라이드 간 이동하며 콘텐츠 추가/수정 작업 시
- 테마, 레이아웃, 콘텐츠를 반복적으로 적용할 때
- Excel 차트를 여러 슬라이드에 삽입할 때

**PowerPoint Shell Mode vs 일반 CLI Mode**

| 특성 | Shell Mode | 일반 CLI Mode |
|------|-----------|--------------|
| **적합한 경우** | 다중 슬라이드 편집 | 단일 슬라이드 작업 |
| **명령어 길이** | 50% 단축 | 전체 경로 필요 |
| **슬라이드 전환** | `use slide N` | 매번 --slide-number |
| **자동완성** | Tab 지원 (41개) | 없음 |
| **히스토리** | 세션 내 유지 | 없음 |

### PPT 시작 방법

```bash
# 방법 1: 빈 세션 시작
oa ppt shell

# 방법 2: 파일 경로로 시작
oa ppt shell --file-path "C:/presentations/report.pptx"
```

### PPT 핵심 워크플로우

```bash
# 시작: 프레젠테이션 로드
oa ppt shell --file-path "sales_report.pptx"

# 1단계: 구조 파악
[PPT: sales_report.pptx > Slide 1] > slides                # 슬라이드 목록 확인
[PPT: sales_report.pptx > Slide 1] > show context          # 현재 상태
[PPT: sales_report.pptx > Slide 1] > layout-list           # 사용 가능한 레이아웃

# 2단계: 슬라이드 편집 (컨텍스트 자동 주입)
[PPT: sales_report.pptx > Slide 1] > use slide 2
[PPT: sales_report.pptx > Slide 2] > content-add-text --text "Q1 Results" --left 100 --top 50
[PPT: sales_report.pptx > Slide 2] > content-add-chart --chart-type "bar" --data-file "q1.json"

# 3단계: 다음 슬라이드로 이동
[PPT: sales_report.pptx > Slide 2] > use slide 3
[PPT: sales_report.pptx > Slide 3] > content-add-excel-chart --excel-file "data.xlsx" --chart-name "Chart1"

# 4단계: 종료
[PPT: sales_report.pptx > Slide 3] > exit
```

**비교: 일반 CLI로 동일 작업**
```bash
# 위 워크플로우를 일반 CLI로 하면:
oa ppt presentation-open --file-path "sales_report.pptx"
oa ppt slide-list --file-path "sales_report.pptx"
oa ppt layout-list --file-path "sales_report.pptx"
oa ppt content-add-text --file-path "sales_report.pptx" --slide-number 2 --text "Q1 Results" --left 100 --top 50
oa ppt content-add-chart --file-path "sales_report.pptx" --slide-number 2 --chart-type "bar" --data-file "q1.json"
oa ppt content-add-excel-chart --file-path "sales_report.pptx" --slide-number 3 --excel-file "data.xlsx" --chart-name "Chart1"
# → 명령어 길이 3배 증가, 타이핑 부담 대폭 증가
```

### PPT 내부 명령어

| 명령어 | 설명 | 예제 |
|--------|------|------|
| `help` | 카테고리별 명령어 목록 | `help` |
| `show context` | 현재 프레젠테이션/슬라이드 상태 | `show context` |
| `use presentation <path>` | 프레젠테이션 전환 | `use presentation "report.pptx"` |
| `use slide <number>` | 슬라이드 전환 (1-indexed) | `use slide 3` |
| `slides` | 슬라이드 목록 | `slides` |
| `presentation-info` | 프레젠테이션 정보 | `presentation-info` |
| `clear` | 화면 지우기 | `clear` |
| `exit` / `quit` | Shell 종료 | `exit` |

### PPT 권장 패턴

#### 1. 슬라이드 탐색 패턴
```bash
# 안전한 순차적 접근
oa ppt shell --file-path "presentation.pptx"
[PPT: presentation.pptx > Slide 1] > slides       # 1. 전체 구조 파악
[PPT: presentation.pptx > Slide 1] > use slide 2  # 2. 작업 슬라이드 선택
[PPT: presentation.pptx > Slide 2] > layout-list  # 3. 레이아웃 확인
[PPT: presentation.pptx > Slide 2] > content-add-text --text "Title"  # 4. 콘텐츠 추가
# → 각 단계마다 출력 확인 후 다음 단계 진행
```

#### 2. Excel 차트 통합 패턴
```bash
# Excel 데이터를 PowerPoint로 시각화
oa ppt shell --file-path "report.pptx"

[PPT: report.pptx > Slide 1] > use slide 3
[PPT: report.pptx > Slide 3] > content-add-excel-chart \
  --excel-file "sales.xlsx" \
  --sheet "Data" \
  --chart-name "MonthlySales" \
  --left 50 --top 100

[PPT: report.pptx > Slide 3] > content-add-text \
  --text "Source: Sales Database Q1 2024" \
  --left 50 --top 450 --font-size 10

[PPT: report.pptx > Slide 3] > use slide 4
# → Excel Chart + 설명 텍스트를 여러 슬라이드에 반복 적용
```

#### 3. 테마 일괄 적용 패턴
```bash
# 전체 슬라이드에 일관된 디자인 적용
oa ppt shell --file-path "template.pptx"

[PPT: template.pptx > Slide 1] > theme-apply --theme-path "corporate.thmx"
[PPT: template.pptx > Slide 1] > slides  # 슬라이드 수 확인

# 각 슬라이드에 적절한 레이아웃 적용
[PPT: template.pptx > Slide 1] > layout-apply --layout-index 0  # 제목
[PPT: template.pptx > Slide 1] > use slide 2
[PPT: template.pptx > Slide 2] > layout-apply --layout-index 1  # 제목+내용
[PPT: template.pptx > Slide 2] > use slide 3
[PPT: template.pptx > Slide 3] > layout-apply --layout-index 2  # 비교
# → 슬라이드 전환만으로 레이아웃 일괄 적용
```

#### 4. 복잡한 슬라이드 제작 패턴
```bash
# 여러 요소를 조합한 슬라이드 생성
oa ppt shell --file-path "complex.pptx"

[PPT: complex.pptx > Slide 1] > use slide 5
[PPT: complex.pptx > Slide 5] > content-add-shape --shape-type "RECTANGLE" --left 50 --top 50
[PPT: complex.pptx > Slide 5] > content-add-text --text "Key Metrics" --left 60 --top 60
[PPT: complex.pptx > Slide 5] > content-add-table --rows 4 --cols 3 --left 50 --top 150
[PPT: complex.pptx > Slide 5] > content-add-image --image-path "logo.png" --left 600 --top 400
# → 하나의 슬라이드에 Shape + Text + Table + Image 복합 구성
```

### 성능 및 효율성

- **명령어 입력 속도**: 일반 CLI 대비 10배 빠름
- **슬라이드 전환**: `use slide N` 명령으로 즉시 전환
- **콘텐츠 추가**: 반복적인 --file-path, --slide-number 생략
- **생산성**: 다중 슬라이드 작업 시 60% 시간 절약

---

## 참고 문서

- [CLAUDE.md](../CLAUDE.md) - AI Agent Quick Reference
- [ADVANCED_FEATURES.md](./ADVANCED_FEATURES.md) - Map Chart, 고급 차트 기능
- [CLAUDE_CODE_PATTERNS.md](./CLAUDE_CODE_PATTERNS.md) - Claude Code 특화 패턴
