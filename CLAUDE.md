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

### Core Operations
- File operations: open, save, close, create workbooks
- Sheet management: add, delete, rename, activate sheets
- Data operations: read/write ranges, table handling with pandas
- Formatting: cell formatting, borders, auto-fit columns
- Advanced: macro execution, chart creation, value finding

### Reference Documentation
Comprehensive xlwings patterns and examples are documented in `specs/xlwings.md`, including:
- Cross-platform considerations (Windows COM vs macOS AppleScript)
- Asynchronous processing patterns
- Resource management and COM object cleanup
- OS-specific limitations and workarounds

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

### Parameter Handling
- All inputs via CLI options: `--option-name value`
- Large text/data via temporary files with auto-cleanup
- File paths as absolute paths in CLI arguments

### Output Processing
- All scripts return structured JSON with version metadata
- AI agents parse raw output and present user-friendly summaries
- Error messages structured for AI interpretation and user explanation

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