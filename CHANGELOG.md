# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [HeadVer](https://github.com/Line-Mode/headver) versioning.

## [Unreleased]

### Fixed
- **Issue #70**: Fixed COM error 0x800401FD in chart-pivot-create command
  - Added smart recovery logic that detects successful chart creation despite COM connection errors
  - Enhanced error messages with context-specific guidance for users and AI agents
  - Improved chart verification with comprehensive data validation
  - JSON responses now include COM error recovery information
  - Related files: `chart_pivot_create.py`, `utils.py`
  - Test coverage: `tests/test_issue_70_com_error.py`

### Added
- Comprehensive test suite for COM error recovery scenarios
- Enhanced COM error message mapping with recovery information
- Auto-recovery functionality for COM error 0x800401FD (CO_E_OBJNOTCONNECTED)

### Changed
- Improved user feedback messages during COM error recovery
- Enhanced chart creation verification logic with better validation
- Updated COM testing documentation with Issue #70 coverage

---

## Previous Releases

### [10.2540.3] - 2024-09-27
- Added auto-position feature to chart-pivot-create command
- Removed MCP support to fix fastmcp warnings
- Fixed GitHub Actions build configuration
- Added keyring dependency management

### [9.2539.33] - Earlier
- Initial Excel and HWP automation features
- Basic CLI structure implementation
- Cross-platform compatibility support

---

## Notes

- **HeadVer Format**: `v{major}.{yearweek}.{build}`
- **Issue Tracking**: All issues tracked on [GitHub Issues](https://github.com/pyhub-ai/pyhub-office-automation/issues)
- **Testing**: Comprehensive test coverage with pytest
- **Platform**: Windows (primary), macOS (limited), Linux (basic)