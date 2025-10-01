# PowerPoint Shell Mode - 기본 워크플로우 예제

이 문서는 PowerPoint Shell Mode의 실제 사용 예제를 단계별로 설명합니다.

## 예제 1: 기본 프레젠테이션 제작

### 시나리오
새로운 프레젠테이션을 만들고 여러 슬라이드에 콘텐츠를 추가합니다.

### 실행 방법

```bash
# Shell 시작
$ oa ppt shell --file-path "sales_report.pptx"

# 1단계: 현재 상태 확인
[PPT: sales_report.pptx > Slide 1] > show context
Current Context:
  Presentation: sales_report.pptx
  Path: C:/presentations/sales_report.pptx
  Total Slides: 5
  Active Slide: 1

# 2단계: 슬라이드 목록 확인
[PPT: sales_report.pptx > Slide 1] > slides
Available slides in sales_report.pptx:
  1. Title Slide (Active)
  2. Content Slide
  3. Data Slide
  4. Chart Slide
  5. Summary Slide

# 3단계: 슬라이드 2로 이동하여 콘텐츠 추가
[PPT: sales_report.pptx > Slide 1] > use slide 2
✓ Active slide: 2/5

[PPT: sales_report.pptx > Slide 2] > content-add-text \
  --text "Q1 Sales Performance" \
  --left 100 --top 50 --width 600 --height 80 \
  --font-size 32 --bold

# 4단계: 이미지 추가
[PPT: sales_report.pptx > Slide 2] > content-add-image \
  --image-path "chart.png" \
  --left 150 --top 200 --width 500 --height 300

# 5단계: 완료 및 저장
[PPT: sales_report.pptx > Slide 2] > presentation-save
✓ Presentation saved

[PPT: sales_report.pptx > Slide 2] > exit
Goodbye!
```

### 학습 포인트
- ✅ `show context`로 현재 상태 파악
- ✅ `slides`로 전체 구조 확인
- ✅ `use slide` 명령으로 빠른 슬라이드 전환
- ✅ 컨텍스트 자동 주입으로 --file-path, --slide-number 생략

---

## 예제 2: Excel 차트를 PowerPoint에 삽입

### 시나리오
Excel 파일의 차트를 PowerPoint 슬라이드에 삽입합니다.

### 실행 방법

```bash
$ oa ppt shell

# 1단계: 프레젠테이션 로드
[PPT: None > Slide None] > use presentation "report.pptx"
✓ Presentation set: report.pptx
✓ Active slide: 1/10

# 2단계: 작업 슬라이드로 이동
[PPT: report.pptx > Slide 1] > use slide 5
✓ Active slide: 5/10

# 3단계: Excel 차트 삽입
[PPT: report.pptx > Slide 5] > content-add-excel-chart \
  --excel-file "sales_data.xlsx" \
  --sheet "Dashboard" \
  --chart-name "MonthlySales" \
  --left 50 --top 100 \
  --width 600 --height 400

# 4단계: 차트 설명 텍스트 추가
[PPT: report.pptx > Slide 5] > content-add-text \
  --text "Source: Sales Database Q1 2024" \
  --left 50 --top 520 \
  --font-size 10 --italic

# 5단계: 다음 슬라이드에도 동일 패턴 적용
[PPT: report.pptx > Slide 5] > use slide 6
[PPT: report.pptx > Slide 6] > content-add-excel-chart \
  --excel-file "sales_data.xlsx" \
  --sheet "Dashboard" \
  --chart-name "RegionalSales" \
  --left 50 --top 100

[PPT: report.pptx > Slide 6] > exit
```

### 학습 포인트
- ✅ Excel 차트를 PowerPoint에 직접 삽입
- ✅ 슬라이드 전환으로 여러 차트 연속 삽입
- ✅ 차트 + 텍스트 조합으로 완성도 높은 슬라이드 제작

---

## 예제 3: 테마 및 레이아웃 일괄 적용

### 시나리오
프레젠테이션에 테마를 적용하고 각 슬라이드에 적절한 레이아웃을 지정합니다.

### 실행 방법

```bash
$ oa ppt shell --file-path "template.pptx"

# 1단계: 사용 가능한 레이아웃 확인
[PPT: template.pptx > Slide 1] > layout-list
Available layouts:
  0: Title Slide
  1: Title and Content
  2: Section Header
  3: Two Content
  4: Comparison
  5: Title Only
  6: Blank

# 2단계: 전체 슬라이드 확인
[PPT: template.pptx > Slide 1] > slides
Available slides: 8 total

# 3단계: 테마 적용 (선택사항)
[PPT: template.pptx > Slide 1] > theme-apply --theme-path "corporate.thmx"
✓ Theme applied successfully

# 4단계: 각 슬라이드에 레이아웃 적용
[PPT: template.pptx > Slide 1] > layout-apply --layout-index 0  # 제목 슬라이드
✓ Layout applied to slide 1

[PPT: template.pptx > Slide 1] > use slide 2
[PPT: template.pptx > Slide 2] > layout-apply --layout-index 1  # 제목+내용
✓ Layout applied to slide 2

[PPT: template.pptx > Slide 2] > use slide 3
[PPT: template.pptx > Slide 3] > layout-apply --layout-index 4  # 비교
✓ Layout applied to slide 3

[PPT: template.pptx > Slide 3] > use slide 4
[PPT: template.pptx > Slide 4] > layout-apply --layout-index 1  # 제목+내용

# 계속 반복...
[PPT: template.pptx > Slide 4] > exit
```

### 학습 포인트
- ✅ `layout-list`로 사용 가능한 레이아웃 먼저 확인
- ✅ 슬라이드 전환 + 레이아웃 적용 반복 패턴
- ✅ 테마 적용으로 일관된 디자인 유지

---

## 예제 4: 복잡한 슬라이드 구성

### 시나리오
하나의 슬라이드에 여러 요소(도형, 텍스트, 표, 이미지)를 조합합니다.

### 실행 방법

```bash
$ oa ppt shell --file-path "dashboard.pptx"

[PPT: dashboard.pptx > Slide 1] > use slide 3

# 1단계: 배경 도형 추가
[PPT: dashboard.pptx > Slide 3] > content-add-shape \
  --shape-type "RECTANGLE" \
  --left 50 --top 50 --width 650 --height 450 \
  --fill-color "EEEEEE"

# 2단계: 제목 텍스트
[PPT: dashboard.pptx > Slide 3] > content-add-text \
  --text "Q1 Performance Dashboard" \
  --left 60 --top 60 --width 600 --height 50 \
  --font-size 28 --bold

# 3단계: 데이터 표 추가
[PPT: dashboard.pptx > Slide 3] > content-add-table \
  --rows 5 --cols 4 \
  --left 60 --top 120 --width 300 --height 150

# 4단계: 차트 이미지 추가
[PPT: dashboard.pptx > Slide 3] > content-add-image \
  --image-path "trend.png" \
  --left 380 --top 120 --width 300 --height 150

# 5단계: 주석 텍스트
[PPT: dashboard.pptx > Slide 3] > content-add-text \
  --text "* Data as of March 31, 2024" \
  --left 60 --top 470 \
  --font-size 10 --italic

[PPT: dashboard.pptx > Slide 3] > exit
```

### 학습 포인트
- ✅ 여러 콘텐츠 요소를 순차적으로 추가
- ✅ 위치와 크기를 정밀하게 지정
- ✅ 도형 → 텍스트 → 표 → 이미지 순서로 레이어링

---

## 예제 5: Tab 자동완성 활용

### 시나리오
명령어를 정확히 기억하지 못할 때 Tab 자동완성을 활용합니다.

### 실행 방법

```bash
$ oa ppt shell

# Tab 키로 명령어 탐색
[PPT: None > Slide None] > pr<TAB>
# 자동완성: presentation-create, presentation-open, presentation-list, presentation-info, presentation-save

[PPT: None > Slide None] > presentation-<TAB>
# 하위 명령 확인

[PPT: None > Slide None] > presentation-list

# 프레젠테이션 로드 후
[PPT: None > Slide None] > use <TAB>
# 자동완성: use presentation, use slide

[PPT: None > Slide None] > use p<TAB>
# 자동완성: use presentation

[PPT: None > Slide None] > use presentation "demo.pptx"

# 슬라이드 작업
[PPT: demo.pptx > Slide 1] > sl<TAB>
# 자동완성: slide-list, slide-add, slide-delete, slide-duplicate, slide-copy, slide-reorder, slides, slideshow-start, slideshow-control

[PPT: demo.pptx > Slide 1] > slides

[PPT: demo.pptx > Slide 1] > co<TAB>
# 자동완성: content-add-text, content-add-image, content-add-shape, content-add-table, content-add-chart, content-add-video, content-add-smartart, content-add-excel-chart, content-add-audio, content-add-equation, content-update

[PPT: demo.pptx > Slide 1] > content-add-text --text "Hello"
```

### 학습 포인트
- ✅ Tab 키로 41개 명령어 모두 탐색 가능
- ✅ 부분 입력 후 Tab으로 자동완성
- ✅ 명령어 오타 방지
- ✅ 명령어를 정확히 몰라도 탐색 가능

---

## Shell Mode vs 일반 CLI 비교

### 동일한 작업을 두 방식으로 비교

**Shell Mode (권장)**:
```bash
$ oa ppt shell --file-path "report.pptx"
[PPT: report.pptx > Slide 1] > use slide 2
[PPT: report.pptx > Slide 2] > content-add-text --text "Title" --left 100 --top 50
[PPT: report.pptx > Slide 2] > content-add-image --image-path "logo.png" --left 100 --top 150
[PPT: report.pptx > Slide 2] > use slide 3
[PPT: report.pptx > Slide 3] > content-add-chart --chart-type "bar" --data-file "data.json"
[PPT: report.pptx > Slide 3] > exit
```
**입력 문자 수**: ~300자

**일반 CLI Mode**:
```bash
$ oa ppt presentation-open --file-path "report.pptx"
$ oa ppt content-add-text --file-path "report.pptx" --slide-number 2 --text "Title" --left 100 --top 50
$ oa ppt content-add-image --file-path "report.pptx" --slide-number 2 --image-path "logo.png" --left 100 --top 150
$ oa ppt content-add-chart --file-path "report.pptx" --slide-number 3 --chart-type "bar" --data-file "data.json"
```
**입력 문자 수**: ~450자

**결과**: Shell Mode가 33% 더 짧음! ✅

---

## 추가 팁

### 1. show context 활용
현재 상태를 주기적으로 확인하여 실수 방지:
```bash
[PPT: report.pptx > Slide 3] > show context
Current Context:
  Presentation: report.pptx
  Path: C:/presentations/report.pptx
  Total Slides: 10
  Active Slide: 3
  All PowerPoint commands will use this context automatically.
```

### 2. clear 명령으로 화면 정리
긴 출력 후 화면 정리:
```bash
[PPT: report.pptx > Slide 3] > clear
```

### 3. help로 명령어 카테고리 확인
```bash
[PPT: report.pptx > Slide 3] > help

Shell Commands (8):
  - help, show, use, clear, exit, quit, slides, presentation-info

PowerPoint Commands by Category:
  Presentation (5): create, open, save, list, info
  Slide (6): list, add, delete, duplicate, copy, reorder
  Content (11): text, image, shape, table, chart, video, smartart, excel-chart, audio, equation, update
  Layout & Theme (4): layout-list, layout-apply, template-apply, theme-apply
  Export (3): pdf, images, notes
  Slideshow (2): start, control
  Other (2): run-macro, animation-add
```

### 4. 명령어 히스토리 활용
- **위/아래 화살표**: 이전 명령 탐색
- **Ctrl+R**: 히스토리 검색 (reverse-i-search)

---

## 문제 해결

### Q: python-pptx 에러가 발생해요
**A**: python-pptx 라이브러리가 설치되어 있는지 확인하세요:
```bash
$ pip install python-pptx
```

### Q: 컨텍스트가 자동 주입되지 않아요
**A**: `show context`로 현재 상태를 확인하세요. 프레젠테이션이나 슬라이드가 설정되지 않았다면 `use` 명령으로 설정하세요.

### Q: 슬라이드 번호가 잘못되었어요
**A**: `slides` 명령으로 전체 슬라이드 목록과 번호를 확인하세요. 슬라이드 번호는 1부터 시작합니다.

### Q: Excel 차트 삽입이 안 돼요
**A**: Excel 파일이 열려있지 않은지 확인하고, 차트 이름이 정확한지 확인하세요.

---

## 다음 단계

- [ ] 복잡한 레이아웃 조합 연습
- [ ] SmartArt 다이어그램 활용
- [ ] 애니메이션 효과 추가
- [ ] 슬라이드쇼 제어 연습
- [ ] 템플릿 기반 프레젠테이션 제작

**PowerPoint Shell Mode를 마스터하여 프레젠테이션 제작 생산성을 10배 향상시키세요!** 🚀
