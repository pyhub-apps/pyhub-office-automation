# Unified Shell Design (Issue #87)

## 개요

Excel Shell과 PowerPoint Shell을 하나의 통합 인터페이스에서 관리하는 Unified Shell Mode를 구현합니다.

## 목표

1. **단일 진입점**: `oa shell` 명령으로 Excel과 PowerPoint 모두 제어
2. **컨텍스트 전환**: Excel ↔ PowerPoint 간 자유로운 전환
3. **상태 유지**: 각 애플리케이션의 상태를 독립적으로 유지
4. **통합 자동완성**: 현재 컨텍스트에 맞는 명령어 자동완성

## 아키텍처

### 1. Context Management

```python
@dataclass
class UnifiedShellContext:
    """Unified shell session context"""
    mode: str = "none"  # "excel", "ppt", "none"

    # Excel context
    excel_workbook_path: Optional[str] = None
    excel_workbook_name: Optional[str] = None
    excel_sheet: Optional[str] = None
    excel_app: Optional[object] = None

    # PowerPoint context
    ppt_presentation_path: Optional[str] = None
    ppt_presentation_name: Optional[str] = None
    ppt_slide_number: Optional[int] = None
    ppt_prs: Optional[object] = None

    def get_prompt_text(self) -> str:
        """Generate prompt text based on current mode"""
        if self.mode == "excel":
            wb_name = self.excel_workbook_name or "None"
            sheet = self.excel_sheet or "None"
            return f"[OA Shell: Excel {wb_name} > Sheet {sheet}] > "
        elif self.mode == "ppt":
            prs_name = self.ppt_presentation_name or "None"
            slide = str(self.ppt_slide_number) if self.ppt_slide_number else "None"
            return f"[OA Shell: PPT {prs_name} > Slide {slide}] > "
        else:
            return "[OA Shell] > "
```

### 2. Command Routing

명령어를 현재 모드에 따라 라우팅합니다:

```python
def execute_unified_command(ctx: UnifiedShellContext, command: str, args: List[str]) -> bool:
    """
    Execute command based on current mode

    Priority:
    1. Unified shell commands (use, switch, show, help, exit, etc.)
    2. Mode-specific commands (excel/ppt commands)
    3. Error if no mode is active
    """

    # Unified shell commands (available in all modes)
    if command in ["help", "show", "clear", "exit", "quit"]:
        return execute_shell_command(command, args)

    # Mode switching commands
    if command == "use":
        return handle_use_command(ctx, args)

    if command == "switch":
        return handle_switch_command(ctx, args)

    # Mode-specific commands
    if ctx.mode == "excel":
        return execute_excel_command(ctx, command, args)
    elif ctx.mode == "ppt":
        return execute_ppt_command(ctx, command, args)
    else:
        console.print("[red]No active mode. Use 'use excel' or 'use ppt' first.[/red]")
        return True
```

### 3. Command Categories

#### Unified Shell Commands (항상 사용 가능)
- `help` - Show all available commands
- `show context` - Display current context (both Excel and PPT)
- `clear` - Clear terminal screen
- `exit` / `quit` - Exit shell
- `use excel <file>` - Switch to Excel mode and open file
- `use ppt <file>` - Switch to PowerPoint mode and open file
- `switch excel` - Switch to Excel mode (if already loaded)
- `switch ppt` - Switch to PowerPoint mode (if already loaded)

#### Excel Commands (Excel 모드에서만)
- 모든 기존 Excel 명령어 (52개)
- `sheets`, `workbook-info`, `range-read`, etc.

#### PowerPoint Commands (PowerPoint 모드에서만)
- 모든 기존 PowerPoint 명령어 (41개)
- `slides`, `presentation-info`, `content-add-text`, etc.

### 4. Tab Autocomplete Strategy

```python
class UnifiedShellCompleter(Completer):
    """Smart autocomplete based on current mode"""

    def get_completions(self, document, complete_event):
        word = document.get_word_before_cursor()

        # Always available: unified commands
        unified_commands = ["help", "show", "clear", "exit", "quit", "use", "switch"]

        # Mode-specific commands
        if ctx.mode == "excel":
            available_commands = unified_commands + EXCEL_COMMANDS
        elif ctx.mode == "ppt":
            available_commands = unified_commands + PPT_COMMANDS
        else:
            available_commands = unified_commands + ["excel", "ppt"]  # for "use excel/ppt"

        for cmd in available_commands:
            if cmd.startswith(word):
                yield Completion(cmd, start_position=-len(word))
```

## 사용 시나리오

### Scenario 1: Excel 작업 후 PowerPoint로 전환

```bash
$ oa shell

# Excel 파일 열기
[OA Shell] > use excel "sales.xlsx"
✓ Excel workbook: sales.xlsx
✓ Active sheet: Data (1/3)

# Excel 작업
[OA Shell: Excel sales.xlsx > Sheet Data] > table-list
[테이블 목록 출력]

[OA Shell: Excel sales.xlsx > Sheet Data] > chart-add --data-range "A1:B10" --chart-type "Column"
✓ Chart created

# PowerPoint로 전환
[OA Shell: Excel sales.xlsx > Sheet Data] > use ppt "report.pptx"
✓ PowerPoint presentation: report.pptx
✓ Active slide: 1/10

# PowerPoint 작업
[OA Shell: PPT report.pptx > Slide 1] > use slide 3
✓ Active slide: 3/10

[OA Shell: PPT report.pptx > Slide 3] > content-add-text --text "Sales Report" --left 100 --top 50
✓ Text added

# 다시 Excel로 돌아가기
[OA Shell: PPT report.pptx > Slide 3] > switch excel
✓ Switched to Excel mode
[OA Shell: Excel sales.xlsx > Sheet Data] >
```

### Scenario 2: 양방향 데이터 연동

```bash
$ oa shell

# Excel 데이터 분석
[OA Shell] > use excel "data.xlsx"
[OA Shell: Excel data.xlsx > Sheet Sheet1] > range-read --range "A1:D100"
[데이터 확인]

# 동일 세션에서 PowerPoint 차트 추가
[OA Shell: Excel data.xlsx > Sheet Sheet1] > use ppt "presentation.pptx"
[OA Shell: PPT presentation.pptx > Slide 1] > content-add-excel-chart \
  --excel-file "data.xlsx" --sheet "Sheet1" --chart-name "Chart1"
✓ Excel chart inserted

# Excel로 돌아가 다른 시트 작업
[OA Shell: PPT presentation.pptx > Slide 1] > switch excel
[OA Shell: Excel data.xlsx > Sheet Sheet1] > use sheet "Summary"
[OA Shell: Excel data.xlsx > Sheet Summary] >
```

### Scenario 3: 컨텍스트 확인

```bash
[OA Shell: Excel data.xlsx > Sheet Data] > show context

Current Context:
  Mode: Excel

  Excel Context:
    Workbook: data.xlsx
    Path: C:/Work/data.xlsx
    Total Sheets: 5
    Active Sheet: Data

  PowerPoint Context:
    Presentation: report.pptx
    Path: C:/Work/report.pptx
    Total Slides: 10
    Active Slide: 3

  Use 'switch ppt' to return to PowerPoint mode.
```

## 구현 세부사항

### 파일 구조

```
pyhub_office_automation/shell/
├── __init__.py
├── excel_shell.py          # 기존 Excel Shell
├── ppt_shell.py            # 기존 PowerPoint Shell
└── unified_shell.py        # 새로운 Unified Shell
```

### 핵심 함수

1. **unified_shell()**: Main REPL loop
2. **execute_unified_command()**: Command routing
3. **handle_use_command()**: Mode switching with file loading
4. **handle_switch_command()**: Mode switching without file loading
5. **show_unified_context()**: Display both Excel and PPT contexts
6. **UnifiedShellCompleter**: Context-aware autocomplete

### 기존 코드 재사용

Excel Shell과 PowerPoint Shell의 명령 실행 로직을 최대한 재사용합니다:

```python
from pyhub_office_automation.shell.excel_shell import (
    execute_excel_command,
    EXCEL_COMMANDS,
)
from pyhub_office_automation.shell.ppt_shell import (
    execute_ppt_command,
    PPT_COMMANDS,
)
```

## 장점

1. **생산성 향상**: 애플리케이션 간 전환 시 Shell 재시작 불필요
2. **통합 워크플로우**: Excel 데이터를 PowerPoint로 바로 연동
3. **컨텍스트 유지**: 각 애플리케이션 상태가 독립적으로 유지됨
4. **일관된 UX**: Excel/PowerPoint Shell과 동일한 명령어 구조

## 제한사항

1. 동시에 하나의 Excel 워크북과 하나의 PowerPoint 프레젠테이션만 활성화 가능
2. HWP는 별도 Shell로 분리 유지 (python-hwp vs pyhwpx 차이)

## 향후 확장

- Issue #88: Batch scripting support
- Issue #89: HWP Shell integration
- Multi-file support (여러 Excel/PPT 파일 동시 관리)

---

**Design Version**: 1.0
**Author**: Claude Code
**Date**: 2025-10-01
