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
  - `oa hwp <command>` for HWP automation
  - `oa info`, `oa install-guide` for package management

### Module Structure
```
pyhub_office_automation/
├── excel/          # xlwings-based Excel automation scripts
├── hwp/            # pyhwpx-based HWP automation scripts
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
- `click`: CLI framework for all command interfaces
- `pandas`: Data processing for Excel operations
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

### Testing Strategy
- **Unit Tests**: `pytest` for individual script functions
- **CLI Tests**: Direct command execution testing with `--help` validation
- **Edge Case Testing**: Handle missing files, wrong paths, program not installed
- **AI Integration Tests**: Verify AI agent can parse outputs and handle errors

### Package Distribution
- **Target**: PyPI distribution as `pyhub-office-automation`
- **Entry Point**: `oa` command registered via `setup.py` entry_points
- **Installation**: `pip install pyhub-office-automation`

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

**Table Operations (2 commands)**
- `table-read` - Read table data into pandas DataFrame
- `table-write` - Write pandas DataFrame as Excel table

### Core Operations
- File operations: open, save, close, create workbooks
- Sheet management: add, delete, rename, activate sheets
- Data operations: read/write ranges, table handling with pandas
- Formatting: cell formatting, borders, auto-fit columns
- Advanced: macro execution, chart creation, value finding

### Workbook Connection Methods (Issue #14)
All Excel commands now support multiple ways to connect to workbooks, eliminating the need to create new Excel instances for each operation:

#### Connection Options
- **`--file-path`**: Traditional file path (existing behavior)
- **`--use-active`**: Connect to currently active workbook
- **`--workbook-name`**: Connect to open workbook by name (e.g., "Sales.xlsx")

#### Usage Examples
```bash
# Traditional file path approach
oa excel range-read --file-path "data.xlsx" --range "A1:C10"

# Use currently active workbook
oa excel range-read --use-active --range "A1:C10"

# Connect to specific open workbook by name
oa excel range-read --workbook-name "Sales.xlsx" --range "A1:C10"

# AI Agent workflow - efficient consecutive operations
oa excel workbook-open --file-path "report.xlsx"
oa excel sheet-add --use-active --name "Results"
oa excel range-write --use-active --range "A1" --data '["Name", "Score"]'
oa excel table-read --use-active --output-file "summary.csv"

# Workbook discovery and information gathering (Issue #16)
oa excel workbook-list --detailed  # List all open workbooks with details
oa excel workbook-info --use-active --include-sheets  # Get active workbook info with sheet details
oa excel workbook-info --workbook-name "Sales.xlsx" --include-sheets --include-properties  # Comprehensive info
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
- `oa excel workbook-list` - Discover all currently open workbooks
- `oa excel workbook-list --detailed` - Get comprehensive list with file info, sheet counts, save status
- `oa excel workbook-info --use-active --include-sheets` - Analyze active workbook structure

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
oa excel workbook-list --detailed

# 2. Choose appropriate workbook and get structure
oa excel workbook-info --workbook-name "Sales.xlsx" --include-sheets

# 3. Perform operations on identified workbook and sheets
oa excel range-read --workbook-name "Sales.xlsx" --sheet "Data" --range "A1:F100"
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