# Shell Mode 테스트 결과

## Issue #86: Shell Mode Testing & Stabilization

### 테스트 환경
- OS: Windows 10/11
- Python: 3.13+
- 테스트 일자: 2025-10-01
- 테스트 파일:
  - Excel: `test_shell.xlsx` (2 sheets, 2 tables)
  - PowerPoint: `test_shell.pptx` (5 slides)

---

## 1. Excel Shell Mode 테스트

### 1.1 파일 생성 테스트
✅ **PASSED**: `test_shell_excel.py` 실행 성공
- TestData 시트: PeopleTable (6행 x 4열)
- SalesData 시트: SalesTable (4행 x 5열)
- 파일 경로: `C:\Work\pyhub-office-automation\test_shell.xlsx`

### 1.2 Shell 시작 테스트

#### Test Case: 파일 경로로 시작
```bash
uv run oa excel shell --file-path "test_shell.xlsx"
```

**예상 동작**:
- Excel 애플리케이션 시작
- 파일 열기 성공
- 첫 번째 활성 시트 설정
- 프롬프트 표시: `[Excel: test_shell.xlsx > Sheet TestData] >`

**테스트 항목**:
- [ ] 파일이 정상적으로 열리는가?
- [ ] 프롬프트가 올바른 컨텍스트를 표시하는가?
- [ ] 워크북명과 시트명이 정확한가?

### 1.3 Shell 명령어 테스트

#### Test 1: `show context`
```
[Excel: test_shell.xlsx > Sheet TestData] > show context
```

**예상 출력**:
```
Current Context:
  Workbook: test_shell.xlsx
  Path: C:/Work/pyhub-office-automation/test_shell.xlsx
  Total Sheets: 2
  Active Sheet: TestData
  All Excel commands will use this context automatically.
```

**테스트 항목**:
- [ ] 워크북 정보가 정확한가?
- [ ] 시트 수가 맞는가? (2개)
- [ ] 활성 시트가 맞는가? (TestData)

#### Test 2: `sheets`
```
[Excel: test_shell.xlsx > Sheet TestData] > sheets
```

**예상 출력**:
```
Available sheets in test_shell.xlsx:
  1. TestData (Active)
  2. SalesData
```

**테스트 항목**:
- [ ] 모든 시트가 나열되는가?
- [ ] 활성 시트 표시가 정확한가?

#### Test 3: `workbook-info`
```
[Excel: test_shell.xlsx > Sheet TestData] > workbook-info
```

**예상 출력**: 워크북 상세 정보 (all details by default)

**테스트 항목**:
- [ ] 워크북 메타데이터가 표시되는가?
- [ ] 시트 정보가 모두 표시되는가?

#### Test 4: `table-list`
```
[Excel: test_shell.xlsx > Sheet TestData] > table-list
```

**예상 출력**:
```
Tables in test_shell.xlsx:
  Sheet: TestData
    - PeopleTable (6 rows, 4 columns)
  Sheet: SalesData
    - SalesTable (4 rows, 5 columns)
```

**테스트 항목**:
- [ ] 모든 테이블이 나열되는가?
- [ ] 테이블 크기가 정확한가?

#### Test 5: `use sheet`
```
[Excel: test_shell.xlsx > Sheet TestData] > use sheet SalesData
```

**예상 출력**:
```
✓ Active sheet: SalesData (2/2)
```

**예상 프롬프트 변경**: `[Excel: test_shell.xlsx > Sheet SalesData] >`

**테스트 항목**:
- [ ] 시트 전환이 성공하는가?
- [ ] 프롬프트가 업데이트되는가?
- [ ] 시트 번호 표시가 맞는가? (2/2)

#### Test 6: `range-read`
```
[Excel: test_shell.xlsx > Sheet SalesData] > range-read --range A1:E4
```

**예상 동작**: SalesData 시트의 A1:E4 범위 데이터 출력

**테스트 항목**:
- [ ] 데이터가 정확하게 읽히는가?
- [ ] --file-path 인자 자동 주입 확인
- [ ] --sheet 인자 자동 주입 확인

#### Test 7: `range-read` with explicit sheet
```
[Excel: test_shell.xlsx > Sheet SalesData] > range-read --sheet TestData --range A1:D3
```

**예상 동작**: TestData 시트의 A1:D3 범위 읽기 (명시적 시트 지정)

**테스트 항목**:
- [ ] 명시적 시트 지정이 우선되는가?
- [ ] 다른 시트의 데이터가 읽히는가?
- [ ] 현재 활성 시트는 변경되지 않는가?

#### Test 8: `help`
```
[Excel: test_shell.xlsx > Sheet SalesData] > help
```

**예상 출력**:
```
Shell Commands (8):
  help, show, use, clear, exit, quit, sheets, workbook-info

Excel Commands by Category:
  Workbook (4): create, open, list, info
  Sheet (4): activate, add, delete, rename
  Range (2): read, write
  ...
```

**테스트 항목**:
- [ ] 모든 명령어 카테고리가 표시되는가?
- [ ] 명령어 수가 정확한가? (52개: 8 shell + 44 Excel)

#### Test 9: Tab 자동완성
```
[Excel: test_shell.xlsx > Sheet SalesData] > ra<TAB>
```

**예상 동작**: `range-read`, `range-write` 자동완성 제안

**테스트 항목**:
- [ ] Tab 키로 자동완성이 동작하는가?
- [ ] 부분 입력으로 필터링이 되는가?

#### Test 10: 명령어 히스토리
**테스트 항목**:
- [ ] 위/아래 화살표로 이전 명령 탐색 가능한가?
- [ ] Ctrl+R로 히스토리 검색 가능한가?

#### Test 11: `clear`
```
[Excel: test_shell.xlsx > Sheet SalesData] > clear
```

**예상 동작**: 터미널 화면 지우기

**테스트 항목**:
- [ ] 화면이 깨끗하게 지워지는가?
- [ ] 프롬프트가 다시 표시되는가?

#### Test 12: `exit` / `quit`
```
[Excel: test_shell.xlsx > Sheet SalesData] > exit
```

**예상 동작**: Shell 종료, "Goodbye!" 메시지

**테스트 항목**:
- [ ] Shell이 정상 종료되는가?
- [ ] Excel 애플리케이션도 함께 종료되는가?
- [ ] 종료 메시지가 표시되는가?

### 1.4 에러 처리 테스트

#### Test Case: 잘못된 시트명
```
[Excel: test_shell.xlsx > Sheet TestData] > use sheet NonExistent
```

**예상 동작**: 에러 메시지 표시, Shell은 계속 실행

**테스트 항목**:
- [ ] 명확한 에러 메시지가 표시되는가?
- [ ] Shell이 종료되지 않고 계속 실행되는가?

#### Test Case: 잘못된 범위
```
[Excel: test_shell.xlsx > Sheet TestData] > range-read --range ZZ999:AA1000
```

**예상 동작**: 에러 메시지 또는 빈 데이터

**테스트 항목**:
- [ ] 에러가 적절히 처리되는가?
- [ ] Shell이 계속 실행 가능한가?

#### Test Case: 잘못된 명령어
```
[Excel: test_shell.xlsx > Sheet TestData] > invalid-command
```

**예상 동작**: "Unknown command" 에러 메시지

**테스트 항목**:
- [ ] 알 수 없는 명령어를 적절히 처리하는가?
- [ ] help 명령 안내가 표시되는가?

---

## 2. PowerPoint Shell Mode 테스트

### 2.1 파일 생성 테스트
✅ **PASSED**: `test_shell_ppt.py` 실행 성공
- 5개 슬라이드 생성 (Title, Content, Blank, Two Content, Section Header)
- 파일 경로: `C:\Work\pyhub-office-automation\test_shell.pptx`

### 2.2 Shell 시작 테스트

#### Test Case: 파일 경로로 시작
```bash
uv run oa ppt shell --file-path "test_shell.pptx"
```

**예상 동작**:
- 프레젠테이션 열기 성공
- 첫 번째 슬라이드 활성화
- 프롬프트 표시: `[PPT: test_shell.pptx > Slide 1] >`

**테스트 항목**:
- [ ] 파일이 정상적으로 열리는가?
- [ ] 프롬프트가 올바른 컨텍스트를 표시하는가?
- [ ] 프레젠테이션명과 슬라이드 번호가 정확한가?

### 2.3 Shell 명령어 테스트

#### Test 1: `show context`
```
[PPT: test_shell.pptx > Slide 1] > show context
```

**예상 출력**:
```
Current Context:
  Presentation: test_shell.pptx
  Path: C:/Work/pyhub-office-automation/test_shell.pptx
  Total Slides: 5
  Active Slide: 1
```

**테스트 항목**:
- [ ] 프레젠테이션 정보가 정확한가?
- [ ] 슬라이드 수가 맞는가? (5개)
- [ ] 활성 슬라이드가 맞는가? (1)

#### Test 2: `slides`
```
[PPT: test_shell.pptx > Slide 1] > slides
```

**예상 출력**:
```
Available slides in test_shell.pptx:
  1. Title Slide (Active)
  2. Title and Content
  3. Blank
  4. Two Content
  5. Section Header
```

**테스트 항목**:
- [ ] 모든 슬라이드가 나열되는가?
- [ ] 슬라이드 레이아웃 정보가 표시되는가?
- [ ] 활성 슬라이드 표시가 정확한가?

#### Test 3: `presentation-info`
```
[PPT: test_shell.pptx > Slide 1] > presentation-info
```

**예상 출력**: 프레젠테이션 상세 정보

**테스트 항목**:
- [ ] 프레젠테이션 메타데이터가 표시되는가?
- [ ] 슬라이드 정보가 모두 표시되는가?

#### Test 4: `use slide`
```
[PPT: test_shell.pptx > Slide 1] > use slide 3
```

**예상 출력**:
```
✓ Active slide: 3/5
```

**예상 프롬프트 변경**: `[PPT: test_shell.pptx > Slide 3] >`

**테스트 항목**:
- [ ] 슬라이드 전환이 성공하는가?
- [ ] 프롬프트가 업데이트되는가?
- [ ] 슬라이드 번호 표시가 맞는가? (3/5)

#### Test 5: `content-add-text`
```
[PPT: test_shell.pptx > Slide 3] > content-add-text --text "Hello Shell" --left 100 --top 100 --width 400 --height 50
```

**예상 동작**: 슬라이드 3에 텍스트 박스 추가 성공

**테스트 항목**:
- [ ] 텍스트가 정상 추가되는가?
- [ ] --file-path 자동 주입 확인
- [ ] --slide-number 자동 주입 확인 (3)

#### Test 6: `content-add-shape`
```
[PPT: test_shell.pptx > Slide 3] > content-add-shape --shape-type RECTANGLE --left 50 --top 200 --width 200 --height 100
```

**예상 동작**: 슬라이드 3에 사각형 도형 추가

**테스트 항목**:
- [ ] 도형이 정상 추가되는가?
- [ ] 위치와 크기가 정확한가?

#### Test 7: `help`
```
[PPT: test_shell.pptx > Slide 3] > help
```

**예상 출력**:
```
Shell Commands (8):
  help, show, use, clear, exit, quit, slides, presentation-info

PowerPoint Commands by Category:
  Presentation (5): create, open, save, list, info
  Slide (6): list, add, delete, duplicate, copy, reorder
  Content (11): text, image, shape, table, chart, video, smartart, excel-chart, audio, equation, update
  ...
```

**테스트 항목**:
- [ ] 모든 명령어 카테고리가 표시되는가?
- [ ] 명령어 수가 정확한가? (41개: 8 shell + 33 PPT)

#### Test 8: Tab 자동완성
```
[PPT: test_shell.pptx > Slide 3] > con<TAB>
```

**예상 동작**: `content-*` 명령어들 자동완성 제안

**테스트 항목**:
- [ ] Tab 키로 자동완성이 동작하는가?
- [ ] 11개 content 명령어가 모두 제안되는가?

#### Test 9: `exit` / `quit`
```
[PPT: test_shell.pptx > Slide 3] > exit
```

**예상 동작**: Shell 종료, "Goodbye!" 메시지

**테스트 항목**:
- [ ] Shell이 정상 종료되는가?
- [ ] 종료 메시지가 표시되는가?

### 2.4 에러 처리 테스트

#### Test Case: 잘못된 슬라이드 번호
```
[PPT: test_shell.pptx > Slide 1] > use slide 99
```

**예상 동작**: 에러 메시지 표시 ("Slide 99 not found")

**테스트 항목**:
- [ ] 명확한 에러 메시지가 표시되는가?
- [ ] Shell이 계속 실행되는가?

#### Test Case: 잘못된 Shape Type
```
[PPT: test_shell.pptx > Slide 3] > content-add-shape --shape-type INVALID --left 0 --top 0
```

**예상 동작**: 에러 메시지 표시

**테스트 항목**:
- [ ] 에러가 적절히 처리되는가?
- [ ] Shell이 계속 실행 가능한가?

---

## 3. 발견된 버그 및 개선 사항

### 버그 목록

#### Bug #1: [제목]
- **심각도**: Critical / High / Medium / Low
- **발견 위치**: [파일:함수:줄번호]
- **재현 단계**:
  1.
  2.
  3.
- **예상 동작**:
- **실제 동작**:
- **수정 방안**:

#### Bug #2: Unicode 인코딩 문제 (FIXED)
- **심각도**: Medium
- **발견 위치**: test_shell_excel.py, test_shell_ppt.py
- **문제**: Windows 터미널 cp949 인코딩으로 ✓, ✗ 문자 출력 실패
- **수정 방안**: UTF-8 인코딩 강제 설정 추가
- **상태**: ✅ FIXED

### 개선 사항

#### Improvement #1: [제목]
- **우선순위**: High / Medium / Low
- **현재 상태**:
- **개선 제안**:
- **예상 효과**:

---

## 4. 성능 테스트

### 4.1 Shell 시작 시간
- **Excel Shell**: [측정 필요] 초
- **PowerPoint Shell**: [측정 필요] 초

### 4.2 명령 실행 시간
- **range-read (100 rows)**: [측정 필요] 초
- **content-add-text**: [측정 필요] 초
- **use sheet/slide**: [측정 필요] 초

### 4.3 메모리 사용량
- **Excel Shell (idle)**: [측정 필요] MB
- **PowerPoint Shell (idle)**: [측정 필요] MB

---

## 5. 테스트 요약

### Excel Shell Mode
- **총 테스트**: 12개
- **통과**: [ ] / 12
- **실패**: [ ] / 12
- **진행 중**: [ ] / 12

### PowerPoint Shell Mode
- **총 테스트**: 9개
- **통과**: [ ] / 9
- **실패**: [ ] / 9
- **진행 중**: [ ] / 9

### 에러 처리
- **총 테스트**: 5개
- **통과**: [ ] / 5
- **실패**: [ ] / 5

### 전체 결과
- **총 테스트**: 26개
- **통과**: [ ] / 26
- **실패**: [ ] / 26
- **통과율**: [ ] %

---

## 6. 다음 단계

### 우선순위 작업
1. [ ] 발견된 버그 수정
2. [ ] 에러 메시지 개선
3. [ ] 성능 최적화
4. [ ] 문서 업데이트

### 추가 테스트 필요 항목
1. [ ] 복잡한 명령어 조합 테스트
2. [ ] 장시간 실행 안정성 테스트
3. [ ] 여러 파일 동시 작업 테스트
4. [ ] 대용량 데이터 처리 테스트

---

## 7. 테스트 실행 방법

### 자동 테스트 파일 생성
```bash
# Excel 테스트 파일 생성
uv run python test_shell_excel.py

# PowerPoint 테스트 파일 생성
uv run python test_shell_ppt.py
```

### Shell 실행
```bash
# Excel Shell
uv run oa excel shell --file-path "test_shell.xlsx"

# PowerPoint Shell
uv run oa ppt shell --file-path "test_shell.pptx"
```

### 테스트 시나리오 실행
각 Shell에서 위 명령어들을 순서대로 실행하고 체크리스트를 확인하세요.

---

**테스트 작성자**: Claude Code
**마지막 업데이트**: 2025-10-01
