# pyhub-office-automation

**AI 에이전트를 위한 Office 자동화 CLI 도구**

Excel과 HWP 문서를 명령줄에서 제어하는 Python 패키지입니다. JSON 출력과 구조화된 에러 처리로 AI 에이전트가 쉽게 사용할 수 있도록 설계되었습니다.

## 🤖 LLM/AI 에이전트를 위한 핵심 기능

- **구조화된 JSON 출력**: 모든 명령어가 AI 파싱에 최적화된 JSON 반환
- **스마트 연결 방법**: 옵션 없이 활성 워크북 자동 선택, `--workbook-name`으로 Excel 재실행 없이 연속 작업
- **컨텍스트 인식**: `workbook-list`로 현재 상황 파악 후 적절한 작업 수행
- **🆕 테이블 구조 즉시 파악**: `table-list`로 컬럼명+샘플데이터를 한 번에 제공, 추가 API 호출 불필요
- **에러 방지**: 작업 전 상태 확인으로 안전한 자동화 워크플로우
- **한국 환경 최적화**: 한글 파일명 지원, HWP 자동화 (Windows)


## 🚀 빠른 시작

### 일반 모드 (개별 명령어)
```bash
# 설치
pip install pyhub-office-automation

# 설치 확인
oa info

# 현재 열린 Excel 파일 확인 (상세 정보 포함)
oa excel workbook-list

# 테이블 구조와 샘플 데이터 즉시 파악 (AI 에이전트 최적화)
oa excel table-list

# 발견한 테이블 데이터 읽기 (완전한 table-driven 워크플로우)
oa excel table-read --table-name "GameData" --limit 100

# 대용량 데이터 처리 (페이징과 샘플링)
oa excel table-read --table-name "GameData" --offset 500 --limit 100
oa excel table-read --table-name "GameData" --limit 50 --sample-mode

# 특정 컬럼만 선택하여 Context 절약
oa excel table-read --table-name "GameData" --columns "게임명,글로벌 판매량" --limit 100

# 기존 방식: 일반 셀 범위 읽기 (Excel Table 외부 데이터용)
oa excel range-read --range "A1:C10"
```

### Interactive Shell Mode (NEW - Issue #85)
연속 작업 시 파일/시트를 매번 지정할 필요가 없습니다!

#### 기본 사용법
```bash
# Shell 시작 (3가지 방법)
oa excel shell                              # 활성 워크북 자동 선택
oa excel shell --file-path "report.xlsx"    # 파일 경로로 시작
oa excel shell --workbook-name "Sales.xlsx" # 열린 파일명으로 시작

# Shell 내부 명령어
[Excel: report.xlsx > Sheet1] > show context      # 현재 상태 확인
[Excel: report.xlsx > Sheet1] > sheets            # 시트 목록
[Excel: report.xlsx > Sheet1] > use sheet Data    # 시트 전환
[Excel: report.xlsx > Data] > range-read --range A1:C10  # 컨텍스트 자동 주입
[Excel: report.xlsx > Data] > table-list          # 테이블 목록
[Excel: report.xlsx > Data] > help                # 카테고리별 명령어 목록
[Excel: report.xlsx > Data] > clear               # 화면 지우기
[Excel: report.xlsx > Data] > exit                # 종료 (quit도 가능)
```

#### 실전 워크플로우 예제

**시나리오 1: 데이터 분석 및 차트 생성**
```bash
oa excel shell --workbook-name "sales.xlsx"

[Excel: sales.xlsx > Sheet1] > use sheet RawData
[Excel: sales.xlsx > RawData] > table-list                    # 테이블 구조 파악
[Excel: sales.xlsx > RawData] > table-read --output-file temp.csv
[Excel: sales.xlsx > RawData] > sheet-add --name "Analysis"   # 분석 시트 생성
[Excel: sales.xlsx > RawData] > use sheet Analysis
[Excel: sales.xlsx > Analysis] > chart-add --data-range "RawData!A1:C10" --chart-type "Column"
[Excel: sales.xlsx > Analysis] > chart-configure --name "Chart1" --title "월별 매출"
[Excel: sales.xlsx > Analysis] > exit
```

**시나리오 2: 피벗테이블 생성 (Windows)**
```bash
oa excel shell

[Excel: None > None] > use workbook "report.xlsx"
[Excel: report.xlsx > None] > use sheet Data
[Excel: report.xlsx > Data] > pivot-create --source-range "A1:F100" --expand table --dest-sheet "Pivot" --dest-range "A1"
[Excel: report.xlsx > Data] > use sheet Pivot
[Excel: report.xlsx > Pivot] > pivot-configure --pivot-name "PivotTable1" --row-fields "Region,Product" --value-fields "Sales:Sum"
[Excel: report.xlsx > Pivot] > pivot-refresh --pivot-name "PivotTable1"
[Excel: report.xlsx > Pivot] > exit
```

**시나리오 3: 다중 시트 데이터 처리**
```bash
oa excel shell --workbook-name "quarterly_report.xlsx"

[Excel: quarterly_report.xlsx > Sheet1] > sheets              # 모든 시트 확인
[Excel: quarterly_report.xlsx > Sheet1] > use sheet Q1
[Excel: quarterly_report.xlsx > Q1] > range-read --range A1:D50 > q1_data.json
[Excel: quarterly_report.xlsx > Q1] > use sheet Q2
[Excel: quarterly_report.xlsx > Q2] > range-read --range A1:D50 > q2_data.json
[Excel: quarterly_report.xlsx > Q2] > use sheet Q3
[Excel: quarterly_report.xlsx > Q3] > range-read --range A1:D50 > q3_data.json
[Excel: quarterly_report.xlsx > Q3] > use sheet Summary
[Excel: quarterly_report.xlsx > Summary] > range-write --range A1 --data '[["Quarter","Total"],["Q1",1000],["Q2",1200],["Q3",1500]]'
[Excel: quarterly_report.xlsx > Summary] > exit
```

**시나리오 4: 탐색 모드 (Tab 자동완성 활용)**
```bash
oa excel shell

[Excel: None > None] > wo<TAB>        # workbook-list 자동완성
[Excel: None > None] > workbook-list   # 열린 파일 확인
[Excel: None > None] > use w<TAB>      # "use workbook" 자동완성
[Excel: None > None] > use workbook "test.xlsx"
[Excel: test.xlsx > None] > sh<TAB>    # sheets 자동완성
[Excel: test.xlsx > None] > sheets     # 시트 목록 확인
[Excel: test.xlsx > None] > use sheet TestData
[Excel: test.xlsx > TestData] > ra<TAB>    # range-read 자동완성
[Excel: test.xlsx > TestData] > range-read --range A1:C10
[Excel: test.xlsx > TestData] > exit
```

**Shell Mode 장점:**
- ✅ **워크북/시트 1회 지정**: 컨텍스트 자동 유지로 반복 입력 불필요
- ✅ **명령어 길이 50% 단축**: `--workbook-name`, `--sheet` 인자 생략
- ✅ **Tab 자동완성**: 52개 명령어 (Shell 8개 + Excel 44개) 지원
- ✅ **명령어 히스토리**: 위/아래 화살표로 이전 명령 재사용
- ✅ **컨텍스트 프롬프트**: `[워크북명 > 시트명] >` 형식으로 현재 상태 표시
- ✅ **대화형 탐색**: 명령 결과를 보고 다음 단계 결정 가능
- ✅ **생산성 향상**: 연속 작업 시 최대 10배 빠른 입력 속도

**지원 명령어:**
- **Shell 전용** (8개): help, show, use, clear, exit, quit, sheets, workbook-info
- **Excel 명령** (44개): 모든 Excel CLI 명령 (Range, Workbook, Sheet, Table, Chart, Pivot 등)

---

### PowerPoint Shell Mode (NEW - Issue #85 Phase 5)
프레젠테이션 작업 시 파일/슬라이드를 매번 지정할 필요가 없습니다!

#### 기본 사용법
```bash
# Shell 시작 (2가지 방법)
oa ppt shell                                    # 새 세션 시작
oa ppt shell --file-path "presentation.pptx"    # 파일 경로로 시작

# Shell 내부 명령어
[PPT: presentation.pptx > Slide 1] > show context          # 현재 상태 확인
[PPT: presentation.pptx > Slide 1] > slides                # 슬라이드 목록
[PPT: presentation.pptx > Slide 1] > use slide 3           # 슬라이드 전환
[PPT: presentation.pptx > Slide 3] > content-add-text --text "Hello" --left 100 --top 100
[PPT: presentation.pptx > Slide 3] > help                  # 카테고리별 명령어 목록
[PPT: presentation.pptx > Slide 3] > exit                  # 종료
```

#### 실전 워크플로우 예제

**시나리오 1: 프레젠테이션 제작**
```bash
oa ppt shell --file-path "sales_report.pptx"

[PPT: sales_report.pptx > Slide 1] > slides                # 전체 슬라이드 확인
[PPT: sales_report.pptx > Slide 1] > slide-add --layout 1  # 새 슬라이드 추가
[PPT: sales_report.pptx > Slide 1] > use slide 2
[PPT: sales_report.pptx > Slide 2] > content-add-text --text "Q1 Results" --left 100 --top 50 --width 600 --height 100
[PPT: sales_report.pptx > Slide 2] > content-add-image --image-path "chart.png" --left 100 --top 200
[PPT: sales_report.pptx > Slide 2] > exit
```

**시나리오 2: Excel 차트 삽입**
```bash
oa ppt shell

[PPT: None > Slide None] > use presentation "report.pptx"
[PPT: report.pptx > Slide 1] > use slide 3
[PPT: report.pptx > Slide 3] > content-add-excel-chart --excel-file "data.xlsx" --sheet "Sheet1" --chart-name "Chart1" --left 50 --top 100
[PPT: report.pptx > Slide 3] > content-add-text --text "Data Source: Q1 Sales" --left 50 --top 400
[PPT: report.pptx > Slide 3] > exit
```

**시나리오 3: 다중 슬라이드 편집**
```bash
oa ppt shell --file-path "training.pptx"

[PPT: training.pptx > Slide 1] > slides                    # 슬라이드 구조 확인
[PPT: training.pptx > Slide 1] > use slide 1
[PPT: training.pptx > Slide 1] > layout-apply --layout-index 0  # 제목 슬라이드
[PPT: training.pptx > Slide 1] > use slide 2
[PPT: training.pptx > Slide 2] > layout-apply --layout-index 1  # 제목 및 내용
[PPT: training.pptx > Slide 2] > content-add-shape --shape-type "RECTANGLE" --left 100 --top 100
[PPT: training.pptx > Slide 2] > use slide 3
[PPT: training.pptx > Slide 3] > content-add-table --rows 5 --cols 3 --left 50 --top 100
[PPT: training.pptx > Slide 3] > exit
```

**시나리오 4: 테마 및 레이아웃 적용**
```bash
oa ppt shell --file-path "presentation.pptx"

[PPT: presentation.pptx > Slide 1] > theme-apply --theme-path "corporate.thmx"
[PPT: presentation.pptx > Slide 1] > layout-list                # 사용 가능한 레이아웃 확인
[PPT: presentation.pptx > Slide 1] > slides                     # 모든 슬라이드 확인
[PPT: presentation.pptx > Slide 1] > use slide 2
[PPT: presentation.pptx > Slide 2] > layout-apply --layout-index 3  # 비교 레이아웃
[PPT: presentation.pptx > Slide 2] > exit
```

**Shell Mode 장점:**
- ✅ **프레젠테이션/슬라이드 1회 지정**: 컨텍스트 자동 유지
- ✅ **명령어 길이 50% 단축**: `--file-path`, `--slide-number` 인자 생략
- ✅ **Tab 자동완성**: 41개 명령어 (Shell 8개 + PPT 33개) 지원
- ✅ **명령어 히스토리**: 위/아래 화살표로 이전 명령 재사용
- ✅ **컨텍스트 프롬프트**: `[프레젠테이션명 > Slide 번호] >` 형식
- ✅ **슬라이드 간 이동**: `use slide` 명령으로 빠른 전환
- ✅ **생산성 향상**: 연속 작업 시 최대 10배 빠른 입력 속도

**지원 명령어:**
- **Shell 전용** (8개): help, show, use, clear, exit, quit, slides, presentation-info
- **PowerPoint 명령** (33개):
  - Presentation (5): create, open, save, list, info
  - Slide (6): list, add, delete, duplicate, copy, reorder
  - Content (11): text, image, shape, table, chart, video, smartart, excel-chart, audio, equation, update
  - Layout & Theme (4): layout-list, layout-apply, template-apply, theme-apply
  - Export (3): pdf, images, notes
  - Slideshow (2): start, control
  - Other (2): run-macro, animation-add

---

### Unified Shell Mode (NEW - Issue #87)
**Excel과 PowerPoint를 하나의 Shell에서 통합 관리!**

#### 기본 사용법
```bash
# Unified Shell 시작
oa shell

# Excel 파일 열기
[OA Shell] > use excel "sales.xlsx"
✓ Excel workbook: sales.xlsx
✓ Mode switched to Excel

# Excel 작업
[OA Shell: Excel sales.xlsx > Sheet Data] > table-list
[테이블 목록 출력]

[OA Shell: Excel sales.xlsx > Sheet Data] > range-read --range A1:B10
[데이터 출력]

# PowerPoint로 전환
[OA Shell: Excel sales.xlsx > Sheet Data] > use ppt "report.pptx"
✓ PowerPoint presentation: report.pptx
✓ Mode switched to PowerPoint

# PowerPoint 작업
[OA Shell: PPT report.pptx > Slide 1] > use slide 3
✓ Active slide: 3/10

[OA Shell: PPT report.pptx > Slide 3] > content-add-text --text "Q1 Results" --left 100 --top 50

# 다시 Excel로 돌아가기
[OA Shell: PPT report.pptx > Slide 3] > switch excel
✓ Switched to Excel mode: sales.xlsx

[OA Shell: Excel sales.xlsx > Sheet Data] > exit
Goodbye from Unified Shell!
```

#### 실전 워크플로우 예제

**시나리오 1: Excel 데이터 분석 → PowerPoint 보고서 생성**
```bash
oa shell

# Excel에서 데이터 분석
[OA Shell] > use excel "quarterly_sales.xlsx"
[OA Shell: Excel quarterly_sales.xlsx > Sheet Sheet1] > use sheet Data
[OA Shell: Excel quarterly_sales.xlsx > Sheet Data] > table-read --output-file analysis.csv
[OA Shell: Excel quarterly_sales.xlsx > Sheet Data] > chart-add --data-range "A1:C20" --chart-type "Column" --title "Quarterly Sales"

# PowerPoint로 전환하여 보고서 작성
[OA Shell: Excel quarterly_sales.xlsx > Sheet Data] > use ppt "quarterly_report.pptx"
[OA Shell: PPT quarterly_report.pptx > Slide 1] > use slide 2
[OA Shell: PPT quarterly_report.pptx > Slide 2] > content-add-excel-chart \
  --excel-file "quarterly_sales.xlsx" --sheet "Data" --chart-name "Chart1" --left 50 --top 100

# Excel로 돌아가 다른 분석
[OA Shell: PPT quarterly_report.pptx > Slide 2] > switch excel
[OA Shell: Excel quarterly_sales.xlsx > Sheet Data] > use sheet Summary
[OA Shell: Excel quarterly_sales.xlsx > Sheet Summary] > range-write --range A1 --data '[["Total Sales", 5000000]]'

# PowerPoint에서 마무리
[OA Shell: Excel quarterly_sales.xlsx > Sheet Summary] > switch ppt
[OA Shell: PPT quarterly_report.pptx > Slide 2] > use slide 3
[OA Shell: PPT quarterly_report.pptx > Slide 3] > content-add-text --text "Total Sales: $5M"
[OA Shell: PPT quarterly_report.pptx > Slide 3] > exit
```

**시나리오 2: 컨텍스트 확인 및 관리**
```bash
oa shell

[OA Shell] > use excel "data.xlsx"
[OA Shell: Excel data.xlsx > Sheet Sheet1] > use ppt "presentation.pptx"
[OA Shell: PPT presentation.pptx > Slide 1] > show context

Current Context:
  Active Mode: PPT

  Excel Context:
    Workbook: data.xlsx
    Path: C:/Work/data.xlsx
    Active Sheet: Sheet1
    (Use 'switch excel' to activate)

  PowerPoint Context:
    Presentation: presentation.pptx
    Path: C:/Work/presentation.pptx
    Active Slide: 1
    (Currently active)

[OA Shell: PPT presentation.pptx > Slide 1] > help
[통합 명령어 목록 표시]
```

**Unified Shell 장점:**
- 🔄 **모드 전환**: Excel ↔ PowerPoint 간 자유로운 전환
- 📦 **컨텍스트 보존**: 각 애플리케이션 상태 독립적 유지
- ⚡ **통합 워크플로우**: 데이터 분석 → 시각화 → 보고서 작성을 단일 세션에서
- 🎯 **컨텍스트 인식**: 현재 모드에 맞는 명령어만 자동완성 제공
- 💡 **직관적 UX**: `use` (파일 열기), `switch` (모드 전환) 명령으로 간단 제어
- 🚀 **생산성 극대화**: 애플리케이션 전환 시 Shell 재시작 불필요

**지원 명령어:**
- **통합 Shell** (8개): help, show context, clear, exit, quit, use excel/ppt, switch excel/ppt
- **Excel 모드**: 모든 Excel 명령어 (52개)
- **PowerPoint 모드**: 모든 PowerPoint 명령어 (41개)

**명령어 비교:**
| 작업 | 일반 CLI | Unified Shell | 절감 |
|------|---------|--------------|------|
| Excel 분석 + PPT 보고서 | 15개 명령 | 8개 명령 | 47% ↓ |
| 앱 전환 | Shell 재시작 필요 | `switch` 1회 | 10초 ↓ |
| 컨텍스트 입력 | 매번 --file-path | 1회만 | 80% ↓ |

---

## 📧 Email 자동화 (NEW)

AI 기반 이메일 생성 및 다중 계정 관리 시스템입니다. Windows Credential Manager를 통한 안전한 자격증명 관리를 지원합니다.

### 빠른 시작
```bash
# 계정 설정 (Gmail, Outlook, Naver 지원)
oa email config --provider gmail --username your@gmail.com

# 계정 목록 확인
oa email accounts

# AI로 이메일 생성 및 발송
oa email send --to recipient@example.com --prompt "프로젝트 진행 상황 보고"

# 특정 계정으로 발송
oa email send --account work --to team@company.com --prompt "회의 일정 변경"
```

### 주요 기능
- 🔐 **안전한 계정 관리**: Windows Credential Manager 연동
- 🤖 **AI 이메일 생성**: 프롬프트로 자동 작성
- 📨 **다중 계정 지원**: 업무/개인 계정 분리
- 🔒 **앱 비밀번호**: OAuth2 없이 간단한 인증

📚 상세 매뉴얼: [docs/email.md](docs/email.md)

## 🤖 AI 코드 어시스턴트 설정

각 AI 에이전트에 최적화된 설정 파일을 자동으로 생성합니다:

### 지원 대상
- **Claude Code** → CLAUDE.md 업데이트/생성
- **Gemini CLI** → GEMINI.md 생성
- **Codex CLI** → AGENTS.md 생성
- **모든 AI** → 전체 파일 생성

### 사용법
```bash
# Claude Code 사용자
oa ai-setup claude

# Gemini CLI 사용자
oa ai-setup gemini

# Codex CLI 사용자
oa ai-setup codex

# 모든 AI 지원 파일 생성
oa ai-setup all
```

### 자동 생성되는 내용
설정 파일에는 다음 내용이 포함됩니다:

- ✅ **`oa` 명령어 사용 가이드**: 기본 사용법과 권장 패턴
- ✅ **Python 경로 자동 탐지**: 설치된 Python 환경 자동 설정
- ✅ **Excel 데이터 차트 제안 템플릿**: 데이터 유형별 차트 추천
- ✅ **에러 처리 및 디버깅 가이드**: 자주 발생하는 문제 해결법
- ✅ **워크북 연결 최적화**: 효율적인 Excel 파일 접근 방법

### Python 환경 자동 탐지
시스템의 Python 설치를 자동으로 감지하여 지침에 포함:
- PATH 환경변수 확인
- 일반적인 설치 경로 스캔 (`anaconda3`, `Programs/Python` 등)
- matplotlib 한글 폰트 설정 (Malgun Gothic, 300dpi)

### 예시 출력
```
✅ AI 에이전트 설정 완료!
- 파일 생성: GEMINI.md
- Python 경로 감지: C:\Users\user\anaconda3\python.exe
- 차트 템플릿 추가: 5개 예시
- 다음 명령으로 확인: cat GEMINI.md

💡 사용 팁: AI 에이전트에서 이 파일을 자동으로 읽도록 설정하세요.
```

> **참고**: 이 기능은 [GitHub Issue #56](https://github.com/pyhub-apps/pyhub-office-automation/issues/56)으로 계획되어 있으며, 향후 업데이트에서 구현될 예정입니다.

## 🧠 AI별 맞춤형 가이드 시스템

각 AI 어시스턴트의 특성에 최적화된 사용 가이드를 제공하는 `llm-guide` 명령어입니다. OpenAI Codex의 "Less is More" 원칙을 적용하여 각 AI가 가장 효율적으로 활용할 수 있는 형태로 정보를 제공합니다.

### 지원 AI 어시스턴트 (5개)

| AI 타입 | 특징 | 가이드 스타일 | 출력 라인수 |
|---------|------|---------------|-------------|
| **default** | 범용 표준 | 균형잡힌 워크플로우 | 15-20줄 |
| **codex** | OpenAI Codex CLI | Less is More 최소주의 | 3-5줄 |
| **claude** | Claude Code | 체계적, 안전성 중심 | 20-30줄 |
| **gemini** | Gemini CLI | 대화형, 시각화 중심 | 15-25줄 |
| **copilot** | GitHub Copilot | 범용 표준 (default와 동일) | 15-20줄 |

### 사용법

```bash
# 기본 사용법 (AI 타입 필수)
oa llm-guide <ai_type> [옵션]

# AI별 최적화 가이드
oa llm-guide default           # 범용 표준 가이드
oa llm-guide codex             # 3-5줄 핵심만 (Less is More)
oa llm-guide claude            # 체계적 4단계 워크플로우
oa llm-guide gemini            # 대화형 제안 및 시각화
oa llm-guide copilot           # IDE 통합 중심

# 상세 모드 및 출력 형식
oa llm-guide claude --verbose  # 상세 가이드 (에러 복구, 고급 팁 포함)
oa llm-guide codex --format text      # 텍스트 형식
oa llm-guide gemini --format markdown # 마크다운 형식
oa llm-guide default --lang en        # 영어 출력 (향후 지원)
```

### Codex "Less is More" 원칙 적용 예시

#### Before (기존 방식)
```bash
oa llm-guide
# → 300+ 줄의 README 전체 출력
```

#### After (Codex 최적화)
```bash
oa llm-guide codex
{
  "cmd": "oa excel [operation] --format json",
  "flow": "workbook-list → table-list → operate",
  "out": "json"
}
# → 단 3줄의 핵심 정보만!
```

### AI별 가이드 특징

#### 🔹 **Codex**: 극도로 간소화
- 불필요한 설명 제거, 핵심 명령어만
- JSON 형식으로 구조화된 최소 정보
- "Less is More" 철학 완전 적용

#### 🔹 **Claude**: 체계적 접근
- 4단계 워크플로우: discover → analyze → plan → execute
- 안전성 원칙과 에러 복구 전략
- 상세 모드에서 컨텍스트 발견 및 스마트 실행 가이드

#### 🔹 **Gemini**: 대화형 상호작용
- 대화 흐름: 인사 → 상황파악 → 분석 → 제안 → 실행
- 데이터 패턴별 스마트 제안 (매출, 시계열, 대용량)
- 시각화 우선순위 및 배치 작업 예시

#### 🔹 **Default/Copilot**: 범용 표준
- 모든 AI가 공통으로 활용할 수 있는 균형잡힌 가이드
- 표준 워크플로우와 연결 방법 제공
- 예제와 모범 사례 포함

### 출력 형식 옵션

- **JSON** (기본): AI 파싱 최적화, 구조화된 데이터
- **Text**: 사람이 읽기 쉬운 일반 텍스트
- **Markdown**: 문서화 및 공유용 마크다운

## 📊 핵심 Excel 명령어

### 상황 파악
```bash
oa excel workbook-list                    # 열린 파일 목록 (상세 정보 포함)
oa excel workbook-info                     # 활성 파일 정보 (모든 정보 포함)
oa excel workbook-info --workbook-name "파일.xlsx"  # 특정 파일 구조 (모든 정보 포함)
```

### 데이터 작업
```bash
# 데이터 읽기/쓰기/변환
oa excel range-read --range "A1:C10"
oa excel range-write --range "A1" --data '["이름", "나이", "부서"]'
oa excel range-convert --range "A1:C10"  # 문자열→숫자 자동 변환

# 형식 변환 상세 옵션
oa excel range-convert --range "A1:Z100" --remove-comma  # "1,234" → 1234
oa excel range-convert --range "B2:B100" --remove-currency  # "₩1,000" → 1000
oa excel range-convert --range "C1:C50" --parse-percent  # "50%" → 0.5
oa excel range-convert --range "D1:D100" --expand table --no-save  # 테이블 전체 변환, 저장 안 함

# 테이블 처리
oa excel table-list                           # 🆕 모든 테이블 구조+샘플 데이터 (AI 최적화)
oa excel table-read --output-file "data.csv"
oa excel table-write --range "A1" --data-file "data.csv"
oa excel table-analyze --table-name "Sales"  # 🆕 특정 테이블 메타데이터 생성
oa excel metadata-generate                    # 🆕 모든 테이블 메타데이터 일괄 생성

# Excel Table 관리 (Windows 전용)
oa excel table-create --range "A1:D100" --table-name "SalesData"  # 범위를 Excel Table로 변환
oa excel table-write --data-file "data.csv" --table-name "AutoTable"  # 데이터 쓰기 + Table 생성
```

### 워크북/시트 관리
```bash
oa excel workbook-create --name "새파일" --save-path "report.xlsx"
oa excel sheet-add --name "결과"
oa excel sheet-activate --name "데이터"
```

### 차트 생성 (두 가지 방식)

#### 정적 차트 - chart-add
```bash
# 일반 데이터 범위에서 차트 생성
# 데이터가 변경되어도 차트는 고정된 범위만 표시
oa excel chart-add --data-range "A1:C10" --chart-type "column" --title "매출 현황"

# 자동 배치로 차트 생성
oa excel chart-add --data-range "A1:C10" --auto-position --chart-type "line" --title "추세 분석"
```

#### 동적 피벗차트 - chart-pivot-create (Windows 전용)
```bash
# 기존 피벗테이블 기반으로 차트 생성
# 피벗테이블 필터/재배치 시 차트도 자동 업데이트
oa excel chart-pivot-create --pivot-name "SalesAnalysis" --chart-type "column" --title "동적 매출 분석"

# 다른 시트에 피벗차트 생성
oa excel chart-pivot-create --pivot-name "ProductSummary" --chart-type "pie" --sheet "Dashboard" --position "B2"
```

💡 **차트 선택 가이드**:
- **chart-add**: 고정 데이터, 간단한 시각화, 일회성 차트, 크로스 플랫폼
- **chart-pivot-create**: 대용량 데이터, 동적 분석, 대시보드용, Windows 전용

### 데이터 분석 및 변환 (피벗 준비)

```bash
# 데이터 구조 분석 (피벗테이블 준비상태 확인)
oa excel data-analyze --range "A1:Z100" --expand "table"
oa excel data-analyze --file-path "report.xlsx" --range "Sheet1!A1:K1000"

# 데이터 변환 (피벗테이블용 형식으로 변환)
# 교차표를 세로 형식으로 변환
oa excel data-transform --source-range "A1:M100" --transform-type "unpivot" --output-sheet "PivotReady"

# 병합된 셀 해제 및 값 채우기
oa excel data-transform --source-range "A1" --expand "table" --transform-type "unmerge"

# 모든 필요한 변환 자동 적용
oa excel data-transform --source-range "Data!A1:K999" --transform-type "auto" --expand "table"

# 다단계 헤더를 단일 헤더로 결합
oa excel data-transform --source-range "A1:J50" --transform-type "flatten-headers" --output-sheet "CleanData"
```

### 피벗테이블 생성

```bash
# 피벗테이블 생성 (2단계 필수)
# 1단계: 빈 피벗테이블 생성
# source-range에 시트명 포함 가능 (예: "Data!A1:D100")
oa excel pivot-create --source-range "Data!A1:D100" --expand "table" --dest-sheet "피벗" --dest-range "F1"

# 2단계: 필드 설정 (반드시 필요)
# 간결한 형식 사용 (권장)
oa excel pivot-configure --pivot-name "PivotTable1" \
  --row-fields "지역,제품" \
  --value-fields "매출:Sum" \
  --clear-existing
```

### 여러 객체 자동 배치 (겹침 방지)
```bash
# 첫 번째 피벗테이블 (수동 위치)
oa excel pivot-create --source-range "A1:D100" --dest-range "F1"

# 두 번째 피벗테이블 (자동 배치)
oa excel pivot-create --source-range "A1:D100" --auto-position

# 세 번째 피벗테이블 (사용자 설정)
oa excel pivot-create --source-range "A1:D100" --auto-position --spacing 3 --preferred-position "bottom"

# 정적 차트 자동 배치
oa excel chart-add --data-range "A1:C10" --auto-position --chart-type "line"

# 피벗차트 자동 배치 (Windows)
oa excel chart-pivot-create --pivot-name "PivotTable1" --chart-type "column" --sheet "Dashboard" --position "H1"

# 겹침 검사 후 생성
oa excel chart-add --data-range "A1:C10" --position "K1" --check-overlap
```

## 🔄 AI 워크플로우 예제

### 1. 스마트 상황 파악 후 작업
```bash
# 1단계: 현재 상황 파악
oa excel workbook-list

# 2단계: AI가 JSON 파싱하여 적절한 연결 방법 선택
# 파일이 열려있으면 --workbook-name, 없으면 --file-path 사용

# 3단계: 연속 작업
oa excel workbook-info --workbook-name "sales.xlsx"  # 모든 정보 자동 포함
oa excel range-read --workbook-name "sales.xlsx" --range "Sheet1!A1:Z100"  # 값과 공식 자동 포함
oa excel chart-add --workbook-name "sales.xlsx" --range "A1:C10"
```

### 2. 연속 데이터 처리 (리소스 효율적)
```bash
# Excel을 한 번만 열고 여러 작업 수행
oa excel workbook-open --file-path "data.xlsx"
oa excel sheet-add --name "분석결과"
oa excel range-write --sheet "분석결과" --range "A1" --data '[...]'
oa excel chart-add --sheet "분석결과" --range "A1:C10"
```

### 3. 스마트 데이터 준비 및 피벗테이블 생성 (AI 지원)
```bash
# 1단계: 데이터 구조 자동 분석 (AI 데이터 패턴 감지)
oa excel data-analyze --range "A1:Z1000" --expand "table"
# → AI가 교차표, 병합셀, 다단계헤더, 소계혼재, 넓은형식 등 5가지 문제 자동 감지
# → 피벗테이블 준비도 평가 및 권장 변환방법 제시

# 2단계: AI 권장사항에 따른 자동 데이터 변환
oa excel data-transform --source-range "A1:Z1000" --transform-type "auto" --expand "table" --output-sheet "PivotReady"
# → AI가 감지한 모든 문제를 올바른 순서로 자동 해결
# → 병합셀 해제 → 소계 제거 → 헤더 정리 → 교차표 변환 순으로 적용

# 3단계: 변환된 데이터로 피벗테이블 생성
oa excel pivot-create --source-range "PivotReady!A1:F5000" --expand "table" --dest-sheet "분석결과" --dest-range "A1"

# 4단계: 필드 배치 (변환된 헤더 사용)
oa excel pivot-configure --pivot-name "PivotTable1" \
  --row-fields "카테고리,제품명" \
  --column-fields "측정항목" \
  --value-fields "값:Sum" \
  --clear-existing

# 5단계: 데이터 새로고침
oa excel pivot-refresh --pivot-name "PivotTable1"
```

### 4. 완전한 피벗테이블 워크플로우 (기존 방식)
```bash
# 1단계: 데이터 확인
oa excel range-read --range "A1:K1"  # 헤더 확인

# 2단계: 피벗테이블 생성
oa excel pivot-create --source-range "Data!A1:K999" --expand "table" --dest-sheet "피벗분석" --dest-range "A1"

# 3단계: 필드 배치 (실제 컬럼명 사용)
oa excel pivot-configure --pivot-name "PivotTable1" \
  --row-fields "카테고리,제품명" \
  --column-fields "분기" \
  --value-fields "매출액:Sum,수량:Count" \
  --filter-fields "지역" \
  --clear-existing

# 4단계: 데이터 새로고침
oa excel pivot-refresh --pivot-name "PivotTable1"
```

### 5. Excel Table 기반 고급 피벗 워크플로우 (Windows 전용)
```bash
# 🎯 향상된 워크플로우: Excel Table → 동적 피벗테이블

# 1단계: 데이터를 Excel Table로 변환 (동적 범위 확장을 위해)
oa excel table-write --data-file "sales.csv" --table-name "SalesData" --table-style "TableStyleMedium5"

# 2단계: Excel Table 확인
oa excel table-list

# 3단계: Excel Table 기반 피벗테이블 생성 (범위 자동 확장!)
oa excel pivot-create --source-range "SalesData" --auto-position --pivot-name "SalesPivot"

# 4단계: 피벗테이블 필드 설정
oa excel pivot-configure --pivot-name "SalesPivot" \
  --row-fields "지역,제품" \
  --value-fields "매출:Sum" \
  --clear-existing

# 💡 장점: 새 데이터 추가 시 피벗테이블 범위가 자동으로 확장됨!
# 기존 범위를 Excel Table로 변환하는 경우:
oa excel table-create --range "A1:F100" --table-name "AnalysisData" --headers
```

### 6. 회계 데이터 정리 자동화
```bash
# 문자열 형식의 회계 데이터를 숫자로 일괄 변환
oa excel range-convert --range "A2:F1000" --expand table --remove-currency --remove-comma
# → "₩1,234,567" → 1234567로 자동 변환
# → 피벗테이블이나 계산에 바로 사용 가능

# 백분율 데이터 변환
oa excel range-convert --range "G1:G100" --parse-percent
# → "15.5%" → 0.155로 변환하여 수식에서 바로 활용

# 괄호형 음수 처리 (회계 양식)
oa excel range-convert --range "H1:H200" --remove-comma
# → "(1,000)" → -1000으로 자동 변환
```

### 7. 에러 방지 패턴
```bash
# 안전한 워크플로우: 확인 → 연결 → 작업
oa excel workbook-list | grep "target.xlsx"  # 파일 열림 확인
# 있으면: --workbook-name 사용, 없으면: --file-path로 열기
oa excel range-read --workbook-name "target.xlsx" --range "A1:C10"
```

## 🤖 AI 지원 데이터 분석 기능

### 스마트 데이터 패턴 감지 (data-analyze)
AI가 Excel 데이터를 자동으로 분석하여 피벗테이블 준비 상태를 평가합니다:

**🔍 자동 감지 패턴**:
- **교차표 형식**: 월/분기가 열로 배치된 형태 감지
- **다단계 헤더**: 중첩된 헤더 구조 인식
- **병합된 셀**: 빈 셀로 인한 데이터 불일치 탐지
- **소계 혼재**: 데이터와 소계가 섞여있는 패턴 분석
- **넓은 형식**: 여러 지표가 열로 나열된 구조 식별

**🎯 지능형 분석 결과**:
- 피벗테이블 준비도 0.0~1.0 점수로 평가
- 감지된 문제별 우선순위 권장사항 제공
- 다음 단계 명령어 자동 제안 (변환 타입 포함)

### 지능형 데이터 변환 (data-transform)
AI가 감지한 문제를 최적 순서로 자동 해결합니다:

**🔄 자동 변환 알고리즘 (auto 모드)**:
1. **병합셀 해제 우선**: 데이터 무결성 확보
2. **소계 제거**: 순수 데이터만 추출
3. **헤더 정리**: 다단계 헤더를 단일 헤더로 결합
4. **교차표 변환**: 피벗테이블용 세로 형식으로 변환

**📊 변환 결과 분석**:
- 변환 전후 데이터 크기 비교 (행/열 변화율)
- 적용된 변환 목록과 순서 보고
- 데이터 확장비/감소비 자동 계산

**💡 AI 활용 내부 구조**:
- **pandas 지능형 활용**: DataFrame 패턴 분석으로 데이터 구조 자동 인식
- **통계적 패턴 매칭**: 공통 Excel 문제 패턴을 학습된 알고리즘으로 탐지
- **컨텍스트 인식**: 헤더명, 데이터 분포, 빈 셀 패턴을 종합적으로 분석
- **최적화된 변환 순서**: 데이터 손실 최소화를 위한 단계별 변환 전략

## ✨ 특별 기능

- **🤖 AI 데이터 분석**: 5가지 일반적인 Excel 데이터 문제를 자동 감지하고 해결방안 제시
- **🔄 지능형 자동 변환**: AI가 감지한 모든 문제를 최적 순서로 자동 해결
- **📊 스마트 형식 변환**: 쉼표, 통화기호(₩,$,€,¥,£), 백분율, 괄호형 음수를 숫자로 자동 변환
- **자동 워크북 선택**: 옵션 없이 활성 워크북 자동 사용으로 Excel 재실행 없이 연속 작업
- **`--workbook-name`**: 파일명으로 직접 접근, 경로 불필요
- **워크북 연결 방법**: 옵션 없음(활성), `--file-path`(파일), `--workbook-name`(이름)
- **🎯 자동 배치**: 피벗테이블과 차트가 겹치지 않게 자동으로 빈 공간 찾아 배치
- **⚠️ 겹침 검사**: 지정된 위치의 충돌 여부를 사전 확인하여 경고 제공
- **JSON 최적화**: 모든 출력이 AI 에이전트 파싱에 최적화
- **한글 파일명 지원**: macOS에서 한글 자소분리 문제 자동 해결
- **39개 Excel 명령어**: 워크북/시트/데이터/차트/피벗/도형/슬라이서 전체 지원

## 📋 명령어 발견

```bash
# 전체 명령어 목록 (JSON)
oa excel list
oa hwp list

# 특정 명령어 도움말
oa excel range-read --help

# AI별 맞춤형 LLM 가이드
oa llm-guide default           # 범용 표준 가이드
oa llm-guide codex             # OpenAI Codex (Less is More 원칙)
oa llm-guide claude --verbose  # Claude Code (체계적 워크플로우)
oa llm-guide gemini            # Gemini CLI (대화형 상호작용)
oa llm-guide copilot           # GitHub Copilot (IDE 통합형)

# 출력 형식 변경
oa llm-guide claude --format markdown
oa llm-guide codex --format text
```

## 🖥️ 지원 플랫폼

- **Windows 10/11**: Excel + HWP 전체 기능

---

**문의**: 파이썬사랑방 이진석 (me@pyhub.kr)