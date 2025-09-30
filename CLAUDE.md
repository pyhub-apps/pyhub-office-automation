# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is `pyhub-office-automation`, a Python-based automation package for Excel (xlwings) and HWP (pyhwpx) office document automation. The project is designed to be used by AI agents (primarily Gemini CLI) to provide conversational automation for non-technical users working with Korean office documents.

### Target Platform
- **Operating System**: Windows 10/11 only
- **Python Version**: Python 3.13+
- **Primary Use Case**: AI agent-driven office automation through CLI commands

## Architecture & Design Principles

### CLI Architecture
The package follows a modular CLI design pattern:
- **Main CLI Command**: `oa` (office automation)
- **Category-based Subcommands**:
  - `oa excel <command>` for Excel automation
  - `oa excel shell` for interactive Excel shell mode (Issue #85)
  - `oa hwp <command>` for HWP automation
  - `oa info`, `oa install-guide` for package management

### Module Structure
```
pyhub_office_automation/
├── excel/          # xlwings-based Excel automation scripts
├── hwp/            # pyhwpx-based HWP automation scripts
├── shell/          # Interactive shell mode (Issue #85)
└── cli/            # Main CLI entry points and help commands
```

### Single Responsibility Design
- Each script has one clear purpose and responsibility
- All scripts use `click` framework for CLI interfaces
- Each script maintains its own version information
- Scripts output structured JSON/YAML for AI parsing

### AI Agent Integration Pattern
1. **Self-Documentation**: Scripts provide comprehensive `--help` and `--version` information
2. **Structured Output**: All results returned as JSON with version metadata
3. **Temporary File Handling**: Large data passed via temporary files, auto-cleaned after use
4. **Error Handling**: Structured error responses that AI agents can interpret and explain to users

## Core Dependencies

### Required Libraries
- `xlwings`: Excel automation (Windows COM, macOS AppleScript)
- `pyhwpx`: HWP document automation (Windows COM only)
- `typer`: CLI framework for all command interfaces
- `pandas`: Data processing for Excel operations
- `prompt-toolkit`: Interactive shell mode with autocomplete (Issue #85)
- `click-repl`: REPL integration for Typer/Click commands (Issue #85)
- `pathlib`, `tempfile`: File system utilities

### Platform-Specific Notes
- **Windows**: Full functionality with COM-based automation
- **macOS**: Limited xlwings support (no HWP support)
  - **한글 경로 처리**: 자동 NFC 정규화로 자소분리 문제 해결
  - **경로 정규화**: 모든 파일 경로에 대해 자동으로 Unicode NFC 형태로 변환
- **Docker**: Excel tools disabled

## Development Commands

Since this is an early-stage project, the following development setup is expected:

### Project Setup
```bash
# Create virtual environment
python -m venv .venv
.venv\Scripts\activate  # Windows
source .venv/bin/activate  # macOS/Linux

# Install dependencies (when available)
pip install -e .

# Install for development
pip install -e .[dev]
```

### Build Scripts
The project includes cross-platform build scripts for creating standalone executables:

**Windows (PowerShell)**:
```powershell
# Basic build (onedir mode)
.\build_windows.ps1

# Single executable with metadata
.\build_windows.ps1 -BuildType onefile -GenerateMetadata

# CI mode (no user interaction)
.\build_windows.ps1 -BuildType onefile -CiMode

# Use existing spec file
.\build_windows.ps1 -UseSpec

# Get help
.\build_windows.ps1 -Help
```

**macOS/Linux (Bash)**:
```bash
# Basic build (onedir mode)
./build_macos.sh

# Single executable with metadata
./build_macos.sh --onefile --metadata

# CI mode (no user interaction)
./build_macos.sh --onefile --ci

# Use existing spec file
./build_macos.sh --use-spec

# Get help
./build_macos.sh --help
```

**Build Features**:
- Automatic dependency exclusion for size optimization (matplotlib, scipy, sklearn, tkinter, IPython, jupyter)
- Build metadata generation with SHA256 checksums
- Cross-platform parameter support
- CI/CD integration ready
- Post-build validation and testing

### Code Quality Scripts
**Windows (PowerShell)**:
```powershell
# Run all checks
.\lint.ps1

# Auto-fix formatting issues
.\lint.ps1 -Fix

# Quick checks only
.\lint.ps1 -Quick

# Verbose output
.\lint.ps1 -Verbose
```

**macOS/Linux (Bash)**:
```bash
# Run all checks
./lint.sh

# Auto-fix formatting issues
./lint.sh --fix

# Quick checks only
./lint.sh --quick

# Verbose output
./lint.sh --verbose
```

### Testing Strategy
- **Unit Tests**: `pytest` for individual script functions
- **CLI Tests**: Direct command execution testing with `--help` validation
- **Edge Case Testing**: Handle missing files, wrong paths, program not installed
- **AI Integration Tests**: Verify AI agent can parse outputs and handle errors

### Package Distribution
- **Target**: PyPI distribution as `pyhub-office-automation`
- **Entry Point**: `oa` command registered via `setup.py` entry_points
- **Installation**: `pip install pyhub-office-automation`

### Release Management (HeadVer Versioning)
**IMPORTANT**: 항상 표준화된 버전 태그 생성 스크립트를 사용하세요.

```bash
# 표준 버전 태그 생성 스크립트 사용
python scripts/create_version_tag.py --auto-increment

# 특정 빌드 번호로 생성
python scripts/create_version_tag.py 19 --message "Fix critical bug"

# 미리보기만 (실제 태그 생성하지 않음)
python scripts/create_version_tag.py --dry-run --auto-increment
```

**HeadVer 형식**: `v{major}.{yearweek}.{build}`
- **major**: `.headver` 파일의 메이저 버전 (예: 10)
- **yearweek**: 년도 뒤 2자리 + ISO 주차 2자리 (예: 2539 = 2025년 39주차)
- **build**: 빌드 번호 (자동 증가 또는 수동 지정)

**GitHub Actions 자동 빌드**:
- `v*` 태그 푸시 시 자동으로 EXE 빌드 및 릴리즈 생성
- 스크립트에서 직접 푸시 여부 선택 가능

## Excel Automation Features (xlwings)

### Command Structure (Updated: Issue #16)
Excel commands are organized by category for better usability:

**Sheet Management (4 commands)**
- `sheet-activate` - Activate a specific sheet
- `sheet-add` - Add new sheet to workbook
- `sheet-delete` - Delete sheet from workbook
- `sheet-rename` - Rename existing sheet

**Workbook Operations (4 commands)**
- `workbook-create` - Create new Excel workbook
- `workbook-open` - Open existing workbook or connect to active one
- `workbook-list` - List all currently open workbooks with basic info
- `workbook-info` - Get detailed information about a specific workbook

**Range Operations (2 commands)**
- `range-read` - Read data from cell ranges
- `range-write` - Write data to cell ranges

**Table Operations (5 commands)**
- `table-read` - Read table data into pandas DataFrame
- `table-write` - Write pandas DataFrame as Excel table
- `table-list` - List all Excel tables with columns and sample data
- `table-analyze` - Analyze specific table and generate metadata
- `metadata-generate` - Batch generate metadata for all tables

**Chart Operations (7 commands)**
- `chart-add` - Create static chart from data range
- `chart-pivot-create` - Create dynamic pivot chart (Windows only)
- `chart-list` - List all charts in worksheet
- `chart-configure` - Modify chart properties
- `chart-position` - Reposition existing chart
- `chart-export` - Export chart as image
- `chart-delete` - Remove chart from worksheet

### Core Operations
- File operations: open, save, close, create workbooks
- Sheet management: add, delete, rename, activate sheets
- Data operations: read/write ranges, table handling with pandas
- Chart operations: static charts, dynamic pivot charts, chart management
- Formatting: cell formatting, borders, auto-fit columns
- Advanced: macro execution, pivot tables, value finding

### Chart Selection Guide for AI Agents

**Use `chart-add` for:**
- Simple data visualization from fixed ranges
- One-time charts and basic reports
- Cross-platform compatibility (Windows/macOS)
- Quick chart generation without pivot tables
- Static presentations and documentation
- **Recommended when `chart-pivot-create` encounters timeout issues**

**Use `chart-pivot-create` for:**
- ~~Dynamic data analysis with filtering~~ (Currently limited due to Issue #42)
- Dashboard creation with interactive elements (use `--skip-pivot-link` option)
- Large datasets requiring pivot table aggregation
- Charts based on pivot table data (static mode available)
- Windows-only environments

**Known Issues (Issue #42):**
- `PivotLayout.PivotTable` assignment causes 2-minute timeout
- Use `--skip-pivot-link` option to bypass pivot connection
- Use `--fallback-to-static` (default: true) for automatic fallback
- Alternative: Use `chart-add` command for reliable chart creation

**Decision Logic:**
1. **Data Size**: Large datasets (>1000 rows) → `chart-add` (due to timeout issues)
2. **Interactivity**: Need filtering/drilling → Use pivot table + `chart-add` separately
3. **Platform**: macOS environment → `chart-add` only
4. **Complexity**: Simple visualization → `chart-add`
5. **Existing Pivot**: Pivot table already exists → `chart-add` with pivot data range

### Workbook Connection Methods (Issue #14)
All Excel commands now support multiple ways to connect to workbooks, eliminating the need to create new Excel instances for each operation:

#### Connection Options
- **옵션 없음**: 활성 워크북 자동 사용 (기본값)
- **`--file-path`**: Traditional file path (existing behavior)
- **`--workbook-name`**: Connect to open workbook by name (e.g., "Sales.xlsx")

#### Usage Examples
```bash
# Traditional file path approach
oa excel range-read --file-path "data.xlsx" --range "A1:C10"

# Use currently active workbook (automatic)
oa excel range-read --range "A1:C10"

# Connect to specific open workbook by name
oa excel range-read --workbook-name "Sales.xlsx" --range "A1:C10"

# AI Agent workflow - efficient consecutive operations
oa excel workbook-open --file-path "report.xlsx"
oa excel sheet-add --name "Results"
oa excel range-write --range "A1" --data '["Name", "Score"]'
oa excel table-read --output-file "summary.csv"

# Workbook discovery and information gathering (Issue #16)
oa excel workbook-list  # List all open workbooks with details
oa excel workbook-info  # Get active workbook info with all details (default)
oa excel workbook-info --workbook-name "Sales.xlsx"  # Comprehensive info (all included by default)
```

#### Benefits for AI Agents
- **Resource Efficiency**: Reuse existing Excel applications instead of creating new ones
- **Workflow Continuity**: Seamless multi-step operations on the same workbook
- **User Experience**: Works naturally with user's existing Excel sessions
- **Performance**: Faster execution by avoiding application startup overhead
- **Context Awareness**: Use `workbook-list` and `workbook-info` to understand current work context
- **Smart Targeting**: Avoid unnecessary file operations by checking what's already open
- **Error Prevention**: Validate workbook existence before attempting operations

#### Validation
- Commands validate that exactly one connection method is specified
- Clear error messages guide users to correct usage patterns
- Backward compatibility maintained - existing scripts continue to work

### Reference Documentation
Comprehensive xlwings patterns and examples are documented in `specs/xlwings.md`, including:
- Cross-platform considerations (Windows COM vs macOS AppleScript)
- Asynchronous processing patterns
- Resource management and COM object cleanup
- OS-specific limitations and workarounds

### macOS 한글 경로 처리
macOS에서 한글 파일명/경로 사용 시 자소분리 현상을 자동으로 해결합니다:

#### 문제 상황
- macOS가 한글을 NFD(자소 분리) 형태로 저장
- "한글.xlsx" → "ㅎㅏㄴㄱㅡㄹ.xlsx" 형태로 분리되어 파일 인식 실패

#### 해결 방법
- 모든 파일 경로에 대해 자동으로 NFC(자소 결합) 정규화 적용
- `normalize_path()` 함수가 모든 Excel 명령어에 통합되어 투명하게 처리
- 사용자는 별도 설정 없이 한글 파일명 자연스럽게 사용 가능

#### 적용 범위
- 모든 `--file-path` 옵션
- 파일 저장 경로 (`--save-path`)
- 데이터 파일 경로 (`--data-file`, `--output-file`)

```bash
# macOS에서 한글 파일명 사용 예제
oa excel range-read --file-path "한글데이터.xlsx" --range "A1:C10"
oa excel workbook-create --save-path "새워크북.xlsx" --name "테스트"
```

## HWP Automation Features (pyhwpx)

### Core Operations
- Document operations: open, save, close, create HWP documents
- Text operations: insert, replace, extract text content
- Formatting: text styling, fonts, colors
- Tables: insert, fill data, extract table content
- Advanced: image insertion, page breaks, document merging

### Reference Documentation
Complete pyhwpx usage patterns documented in `specs/pyhwpx.md`, covering:
- Document lifecycle management
- Text and formatting operations
- Table and image handling
- PDF and format conversion capabilities
- Mail merge and template processing

## AI Agent Interaction Patterns

### Command Discovery
AI agents should use these commands to understand available functionality:
- `oa excel list` - List all Excel automation commands
- `oa hwp list` - List all HWP automation commands
- `oa get-help <category> <command>` - Get detailed help for specific commands
- `oa info` - Package version and dependency status

### Context Discovery
AI agents should use these commands to understand current work context:
- `oa excel workbook-list` - Discover all currently open workbooks (comprehensive info by default)
- `oa excel workbook-list` - Get comprehensive list with file info, sheet counts, save status
- `oa excel workbook-info` - Analyze active workbook structure (all details by default)
- `oa excel table-list` - **Enhanced**: List all Excel tables with complete structure, columns, and sample data for immediate context understanding

### Parameter Handling
- All inputs via CLI options: `--option-name value`
- Large text/data via temporary files with auto-cleanup
- File paths as absolute paths in CLI arguments

### Output Processing
- All scripts return structured JSON with version metadata
- AI agents parse raw output and present user-friendly summaries
- Error messages structured for AI interpretation and user explanation

### AI Agent Workflow Examples

#### Context-Aware Data Analysis
```bash
# 1. Discover current work environment
oa excel workbook-list

# 2. Choose appropriate workbook and get structure
oa excel workbook-info --workbook-name "Sales.xlsx"  # All details included by default

# 3. Perform operations on identified workbook and sheets
oa excel range-read --workbook-name "Sales.xlsx" --sheet "Data" --range "A1:F100"
```

#### Enhanced Table-Driven Analysis (New)
```bash
# 1. Get complete table overview with structure and sample data
oa excel table-list

# Response provides immediate insights:
# - Table names and locations
# - Column structures (all columns shown)
# - Sample data (top 5 rows with 50-char limit per cell)
# - Data types and business context
# - No additional API calls needed for basic analysis

# 2. AI agent can now suggest analysis without further data exploration:
# - "I see GameData table with sales columns - shall I create regional sales charts?"
# - "The table has 11 columns including genre and platform - want genre analysis?"
# - "998 rows of game sales data detected - ready for top performers analysis?"

# 3. Proceed directly with targeted analysis based on discovered structure
oa excel chart-add --sheet "Data" --data-range "GameData[글로벌 판매량]" --chart-type "Column"
```

#### Multi-Workbook Analysis
```bash
# 1. List all open workbooks to understand scope
oa excel workbook-list

# 2. Analyze each workbook for unsaved changes
oa excel workbook-info --workbook-name "Report1.xlsx"
oa excel workbook-info --workbook-name "Report2.xlsx"

# 3. Save any unsaved workbooks before proceeding
# (Implementation for save commands to be added)
```

#### Error Prevention Workflow
```bash
# 1. Check if target workbook is already open
oa excel workbook-list | grep "target.xlsx"

# 2. If open, use existing; if not, open new
# Open: oa excel workbook-info --workbook-name "target.xlsx"
# Not open: oa excel workbook-open --file-path "/path/to/target.xlsx"

# 3. Proceed with operations using appropriate connection method
oa excel range-read --workbook-name "target.xlsx" --range "A1:C10"
```

### Installation Guidance
- `oa install-guide` provides step-by-step installation instructions
- AI agents should verify installation before attempting operations
- Guide users through Python setup and package installation process

## Security & Data Handling

### Privacy Protection
- **Critical**: Document content must never be used for AI training
- Temporary files immediately deleted after processing
- Local-only processing, no data transmission to external services

### File Safety
- Validate file paths and prevent directory traversal
- Handle missing programs (HWP not installed, Excel unavailable)
- Graceful error handling for file access issues

## Standards Compliance

The project references Korean government database standardization guidelines in `specs/공공기관_데이터베이스_표준화_지침.md` for:
- Data format standards
- Database naming conventions
- Compliance requirements for government sector usage

When working with this codebase, prioritize:
1. Maintaining the modular, single-responsibility design
2. Ensuring AI agent compatibility through structured outputs
3. Following the CLI design patterns established in the PRD
4. Implementing comprehensive error handling for edge cases
5. Maintaining security and privacy standards for document processing

# Code Assistant Context

## oa : pyhub-office-automation CLI utility

+ `oa` 명령을 통해, 현재 구동 중인 엑셀 프로그램과 통신하며 시트 데이터 읽고 쓰기, 피벗 테이블 생성, 차트 생성 등을 할 수 있어.
    - 엑셀 파일 접근에는 `oa` 프로그램을 사용하고, 한 번에 10개 이상의 많은 엑셀 파일을 읽어야할 때에는 효율성을 위해 python과 python 엑셀 라이브러리를 통해 읽어줘.
    - 엑셀 파일을 열기 전에, 반드시 `oa excel workbook-list` 명령으로 열려진 엑셀파일이 있는 지 확인해줘.
    - 파일을 읽을 수 없다면 유저에게 파일 경로를 꼼꼼하게 확인해보라고 알려줘.
+ **ALWAYS** `oa excel --help` 명령으로 지원 명령을 먼저 확인하고, `oa excel 명령 --help` 명령으로 사용법을 확인한 뒤에 명령을 사용해줘.
+ `oa llm-guide` 명령으로 지침을 조회해줘.
+ `--workbook-name` 인자나 `--file-path` 인자를 지정하지 않으면 활성화된 워크북을 참조하고, `--sheet` 인자를 지정하지 않으면, 활성화된 시트를 참조함.
    - 모든 `oa` 명령에서 명시적으로 `--sheet` 인자로 시트명을 지정하여 읽어오자.

## 핵심 사용 패턴

### 1. 작업 전 상황 파악
```bash
# 현재 열린 워크북 확인
oa excel workbook-list

# 활성 워크북 정보 확인
oa excel workbook-info  # All details included by default
```

### 2. 워크북 연결 방법
- **자동 연결**: 옵션 없이 사용하면 활성 워크북 자동 사용 (기본값)
- **파일 경로**: `--file-path "경로/파일명.xlsx"`
- **워크북 이름**: `--workbook-name "파일명.xlsx"`

### 3. 데이터 읽기/쓰기
```bash
# 범위 데이터 읽기
oa excel range-read --sheet "Sheet1" --range "A1:C10"

# 데이터 쓰기
oa excel range-write --sheet "Sheet1" --range "A1" --data '[["Name", "Score"], ["Alice", 95]]'

# 테이블 읽기 (pandas DataFrame으로)
oa excel table-read --sheet "Sheet1" --output-file "data.csv"
```

### 4. 차트 생성
```bash
# 기본 차트 생성
oa excel chart-add --sheet "Sheet1" --data-range "A1:B10" --chart-type "Column" --title "Sales Chart"

# 피벗 차트 생성 (Windows만)
oa excel chart-pivot-create --sheet "Sheet1" --data-range "A1:D100" --rows "Category" --values "Sales"
```

## 에러 방지 워크플로우

1. **항상 workbook-list로 시작**: 현재 상황 파악
2. **명시적 시트 지정**: `--sheet` 옵션 사용
3. **단계별 진행**: 복잡한 작업을 작은 단위로 분할
4. **경로 확인**: 파일 경로는 절대 경로나 정확한 상대 경로 사용

## python 직접 실행

+ oa 가 지원하는 기능 외에 추가로 데이터 변환/분석이 필요하면 아래 경로의 python을 활용해.
    - python 경로 : C:\Users\pyhub\AppData\Local\Programs\Python\Python313\python.EXE
+ 설치되지 않은 라이브러리는 `C:\Users\pyhub\AppData\Local\Programs\Python\Python313\python.EXE -m pip install 팩키지명` 명령으로 설치해
+ matplotlib 차트 생성에서는 Malgun Gothic 폰트를 사용하고, 300dpi 로 생성하자.

### Python 사용 예시

```python
# 한글 폰트 설정 (matplotlib)
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm

# 한글 폰트 설정
plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False

# 고해상도 설정
plt.rcParams['figure.dpi'] = 300
plt.rcParams['savefig.dpi'] = 300
```

### 대용량 데이터 처리

```python
# 여러 Excel 파일 일괄 처리 (10개 이상 파일 시)
import pandas as pd
from pathlib import Path

def process_multiple_files(file_pattern):
    all_data = []
    for file_path in Path().glob(file_pattern):
        df = pd.read_excel(file_path)
        df['source_file'] = file_path.name
        all_data.append(df)

    return pd.concat(all_data, ignore_index=True)

# 사용 예시
combined_data = process_multiple_files("data/*.xlsx")
```

### 추천 라이브러리

- **pandas**: Excel/CSV 데이터 처리
- **openpyxl**: Excel 파일 읽기/쓰기
- **matplotlib**: 차트 생성
- **seaborn**: 통계 차트
- **numpy**: 수치 계산

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

## Claude Code 특화 기능

### 상세 분석 및 체계적 접근

Claude의 깊이 있는 분석 능력을 활용한 Excel 자동화 패턴:

#### 코드 품질 및 구조 분석
```python
# Claude Code가 excel 자동화 스크립트를 분석할 때 중점 사항
def analyze_excel_workflow():
    """
    1. 데이터 무결성 검증
    2. 에러 처리 패턴
    3. 성능 최적화 기회
    4. 코드 재사용성
    """

    # 단계별 검증 워크플로우
    steps = [
        "oa excel workbook-list",  # 현황 파악
        "데이터 구조 분석",                    # 스키마 검토
        "비즈니스 로직 검증",                  # 요구사항 부합성
        "성능 및 확장성 검토"                  # 최적화 기회
    ]

    return steps
```

### 문제 해결 방법론

#### 체계적 디버깅 접근
```bash
# 1. 상황 진단
oa excel workbook-list --format json

# 2. 데이터 구조 분석
oa excel workbook-info  # All details included by default

# 3. 샘플 데이터 검증
oa excel range-read --sheet "Sheet1" --range "A1:E5"

# 4. 에러 재현 및 분석
# (문제가 되는 명령어 단계별 실행)

# 5. 해결책 구현 및 검증
```

### 코드 리뷰 및 최적화

#### Excel 자동화 코드 리뷰 체크리스트
```python
def review_excel_automation():
    """
    Claude Code의 Excel 자동화 코드 리뷰 포인트
    """
    checklist = {
        "에러 처리": [
            "파일 존재 여부 확인",
            "시트 존재 여부 확인",
            "범위 유효성 검증",
            "데이터 타입 검증"
        ],
        "성능": [
            "대용량 데이터 처리 최적화",
            "메모리 사용량 관리",
            "I/O 작업 최소화",
            "배치 처리 활용"
        ],
        "유지보수성": [
            "모듈화된 함수 설계",
            "설정값 외부화",
            "로깅 및 모니터링",
            "문서화 완성도"
        ]
    }
    return checklist
```

### 고급 Excel 활용 패턴

#### 복합 데이터 분석 파이프라인
```python
import subprocess
import json
import pandas as pd
from pathlib import Path

class ExcelAnalysisPipeline:
    """체계적인 Excel 데이터 분석 파이프라인"""

    def __init__(self, workbook_name=None):
        self.workbook_name = workbook_name
        self.context = {}

    def analyze_structure(self):
        """데이터 구조 분석"""
        cmd = ['oa', 'excel', 'workbook-info']  # All details included by default
        if self.workbook_name:
            cmd.extend(['--workbook-name', self.workbook_name])

        result = subprocess.run(cmd, capture_output=True, text=True)
        self.context['structure'] = json.loads(result.stdout)
        return self.context['structure']

    def extract_data(self, sheet, range_addr):
        """데이터 추출 및 검증"""
        cmd = ['oa', 'excel', 'range-read',
               '--sheet', sheet, '--range', range_addr, '--format', 'json']
        if self.workbook_name:
            cmd.extend(['--workbook-name', self.workbook_name])

        result = subprocess.run(cmd, capture_output=True, text=True)
        data = json.loads(result.stdout)

        # 데이터 품질 검증
        df = pd.DataFrame(data.get('data', []))
        self.context['data_quality'] = {
            'rows': len(df),
            'columns': len(df.columns),
            'null_count': df.isnull().sum().sum(),
            'duplicates': df.duplicated().sum()
        }

        return df

    def generate_insights(self, df):
        """데이터 인사이트 생성"""
        insights = {
            'summary_stats': df.describe().to_dict(),
            'data_types': df.dtypes.to_dict(),
            'missing_data': df.isnull().sum().to_dict()
        }

        # 비즈니스 인사이트 추가
        if 'sales' in df.columns or '매출' in df.columns:
            sales_col = 'sales' if 'sales' in df.columns else '매출'
            insights['sales_analysis'] = {
                'total_sales': df[sales_col].sum(),
                'avg_sales': df[sales_col].mean(),
                'top_performers': df.nlargest(5, sales_col).to_dict()
            }

        return insights

    def create_dashboard(self, insights):
        """대시보드 차트 생성"""
        charts_created = []

        # 요약 통계 차트
        summary_chart = self._create_summary_chart()
        if summary_chart:
            charts_created.append(summary_chart)

        # 추세 분석 차트
        trend_chart = self._create_trend_chart()
        if trend_chart:
            charts_created.append(trend_chart)

        return charts_created

    def _create_summary_chart(self):
        """요약 차트 생성"""
        cmd = ['oa', 'excel', 'chart-add',
               '--sheet', 'Dashboard',
               '--data-range', 'A1:B10',
               '--chart-type', 'Column',
               '--title', '데이터 요약']

        result = subprocess.run(cmd, capture_output=True, text=True)
        return json.loads(result.stdout) if result.returncode == 0 else None
```

### 문서화 및 지식 관리

#### 자동 문서 생성
```python
def generate_analysis_report(pipeline_results):
    """분석 결과 자동 문서화"""
    report = f"""
# Excel 데이터 분석 보고서

## 데이터 개요
- 워크북: {pipeline_results['workbook']}
- 시트 수: {len(pipeline_results['sheets'])}
- 총 데이터 행: {pipeline_results['total_rows']}

## 데이터 품질 평가
- 결측값: {pipeline_results['missing_values']}%
- 중복값: {pipeline_results['duplicates']}개
- 데이터 완성도: {pipeline_results['completeness']}%

## 주요 인사이트
{pipeline_results['insights']}

## 권장 액션
{pipeline_results['recommendations']}

## 생성된 차트
{pipeline_results['charts']}
"""
    return report
```

### Claude Code 장점 활용

1. **정확한 분석**: 데이터 무결성과 비즈니스 로직 검증
2. **체계적 접근**: 단계별 분석 프로세스 설계
3. **품질 관리**: 코드 리뷰와 최적화 제안
4. **지식 정리**: 자동 문서화와 인사이트 요약

#### 🔥 Claude Code + table-list 최적 활용법

**즉시 분석 패턴**:
```bash
# Claude Code가 선호하는 효율적 워크플로우
oa excel table-list --format json
# ☝️ 한 번의 호출로 Claude가 즉시 파악:
# - 테이블 구조 (11개 컬럼: 순위, 게임명, 플랫폼, 발행일, 장르, 퍼블리셔, 판매량x4, 글로벌판매량)
# - 샘플 데이터 (Wii 스포츠 82.74M, 슈퍼 마리오 40.24M 등)
# - 데이터 품질 (998행, 정형화된 숫자 데이터)
# - 비즈니스 컨텍스트 (게임 판매 분석 데이터)

# Claude가 즉시 제안 가능한 분석들:
# 1. "글로벌 판매량 Top 10 막대 차트를 만들어드릴까요?"
# 2. "지역별 판매량 비교 (북미 vs 유럽 vs 일본 vs 기타)는 어떨까요?"
# 3. "장르별 집계나 플랫폼별 분석도 가능합니다."
# 4. "발행 연도별 트렌드 분석도 해볼까요?"
```

**Smart Chart Recommendation Engine**:
```python
def claude_smart_chart_suggestions(table_data):
    """
    Claude Code가 table-list 데이터를 분석해 최적 차트 추천
    """
    recommendations = []

    # 컬럼 분석 기반 자동 추천
    columns = table_data.get("columns", [])
    sample_data = table_data.get("sample_data", [])

    if "글로벌 판매량" in columns and "게임명" in columns:
        recommendations.append({
            "type": "Column",
            "title": "게임별 글로벌 판매량 Top 10",
            "reason": "순위 데이터와 판매량 수치로 Top 10 시각화 최적",
            "command": "oa excel chart-add --data-range 'GameData[글로벌 판매량]' --chart-type 'Column'"
        })

    if "북미 판매량" in columns and "유럽 판매량" in columns:
        recommendations.append({
            "type": "Scatter",
            "title": "북미 vs 유럽 판매량 상관관계",
            "reason": "두 지역 판매량 간의 상관성 분석",
            "command": "oa excel chart-add --x-range 'GameData[북미 판매량]' --y-range 'GameData[유럽 판매량]'"
        })

    return recommendations

# 실제 활용: Claude가 즉시 적절한 차트 제안
chart_suggestions = claude_smart_chart_suggestions(table_list_response["data"]["tables"][0])
```

**Data Quality Instant Assessment**:
```python
def claude_data_quality_check(sample_data):
    """
    샘플 데이터만으로 Claude가 즉시 품질 평가
    """
    quality_report = {
        "data_completeness": "✅ NULL 값 없음",
        "data_types": "✅ 숫자 데이터 정상 (41.49, 29.02 등)",
        "business_logic": "✅ 판매량 합계 로직 확인 가능 (지역별 → 글로벌)",
        "recommendations": [
            "발행일을 연도 형식으로 변환하여 시계열 분석",
            "판매량 단위 백만장으로 해석하여 차트 레이블링",
            "상위 게임들의 플랫폼 트렌드 분석 가능"
        ]
    }
    return quality_report
```

### 권장 작업 순서

1. **요구사항 분석**: 비즈니스 목표와 데이터 요구사항 명확화
2. **환경 검증**: `oa excel workbook-list`로 현재 상태 확인
3. **데이터 탐색**: 구조 분석 및 샘플 데이터 검토
4. **분석 설계**: 단계별 분석 프로세스 설계
5. **실행 및 검증**: 각 단계별 결과 검증
6. **결과 정리**: 인사이트 요약 및 액션 아이템 제시


---

## 설정 파일 정보

- **생성 대상**: Claude Code
- **생성 시간**: 2025-09-24 00:05:37
- **패키지 버전**: 9.2539.33
- **Python 탐지**: 활성화

이 파일은 `oa ai-setup claude` 명령으로 생성되었습니다.
