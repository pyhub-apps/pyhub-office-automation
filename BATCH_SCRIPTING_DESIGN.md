# Batch Execution & Scripting Design (Issue #88)

## 개요

Shell 명령어들을 스크립트 파일로 저장하고 일괄 실행하는 기능을 구현합니다.

## 목표

1. **반복 작업 자동화**: 매일/매주 반복되는 Shell 작업을 스크립트로 저장
2. **CI/CD 통합**: 자동화 파이프라인에서 Office 문서 처리
3. **재현 가능한 워크플로우**: 작업 과정을 스크립트로 문서화
4. **에러 처리**: 실패 시 로깅 및 재시도 전략

## 스크립트 파일 포맷

### 1. Basic Format (.oas - Office Automation Script)

```bash
# example.oas - Excel 데이터 분석 및 차트 생성
# Comments start with #

# Excel 워크북 열기
use excel "sales_data.xlsx"

# 시트 전환
use sheet "RawData"

# 데이터 읽기
table-list
range-read --range "A1:F100" --output-file "temp_data.csv"

# 분석 시트 생성
sheet-add --name "Analysis"
use sheet "Analysis"

# 차트 생성
chart-add --data-range "RawData!A1:C20" --chart-type "Column" --title "Monthly Sales"

# 저장 및 종료
# (자동으로 저장됨)
```

### 2. Extended Format with Variables

```bash
# monthly_report.oas - 변수를 사용한 월간 보고서
# Variables: ${VAR_NAME} or $VAR_NAME

# 변수 정의
@set REPORT_MONTH = "2024-01"
@set DATA_FILE = "sales_${REPORT_MONTH}.xlsx"
@set OUTPUT_PPT = "report_${REPORT_MONTH}.pptx"

# Excel 데이터 처리
use excel "${DATA_FILE}"
use sheet "Data"
table-read --output-file "analysis.csv"

# PowerPoint 보고서 생성
use ppt "${OUTPUT_PPT}"
use slide 1
content-add-text --text "Monthly Report - ${REPORT_MONTH}" --left 100 --top 50
content-add-excel-chart --excel-file "${DATA_FILE}" --sheet "Data" --chart-name "Chart1"
```

### 3. Advanced Format with Control Flow

```bash
# advanced_workflow.oas - 제어 흐름 포함

@set DATA_DIR = "C:/Work/data"
@set FILES = ["Q1.xlsx", "Q2.xlsx", "Q3.xlsx", "Q4.xlsx"]

# 조건부 실행
@if exists("${DATA_DIR}/Q4.xlsx")
  use excel "${DATA_DIR}/Q4.xlsx"
  table-list
@else
  @echo "Q4 data not found, skipping..."
@endif

# 반복 실행
@foreach file in ${FILES}
  @echo "Processing ${file}..."
  use excel "${DATA_DIR}/${file}"
  use sheet "Data"
  range-read --range "A1:C100"
@endforeach

# 에러 처리
@try
  chart-add --data-range "A1:B10" --chart-type "Column"
@catch
  @echo "Chart creation failed, using default..."
  chart-add --data-range "A1:B10" --chart-type "Line"
@endtry
```

## 실행 방법

### 1. 직접 실행

```bash
# 단일 스크립트 실행
oa batch run workflow.oas

# 변수 오버라이드
oa batch run monthly_report.oas --set REPORT_MONTH="2024-02"

# Dry-run (실행하지 않고 확인만)
oa batch run workflow.oas --dry-run

# Verbose 모드
oa batch run workflow.oas --verbose
```

### 2. Shell에서 실행

```bash
# Unified Shell에서 스크립트 실행
oa shell

[OA Shell] > @run workflow.oas
[OA Shell] > @run monthly_report.oas --set REPORT_MONTH="2024-03"
```

### 3. CI/CD 통합

```yaml
# GitHub Actions 예제
name: Generate Monthly Reports
on:
  schedule:
    - cron: '0 0 1 * *'  # 매월 1일

jobs:
  generate-reports:
    runs-on: windows-latest
    steps:
      - uses: actions/checkout@v3
      - name: Setup Python
        uses: actions/setup-python@v4
        with:
          python-version: '3.13'
      - name: Install dependencies
        run: pip install pyhub-office-automation
      - name: Run batch script
        run: |
          oa batch run monthly_report.oas \
            --set REPORT_MONTH="$(date +%Y-%m)" \
            --log-file report.log
      - name: Upload results
        uses: actions/upload-artifact@v3
        with:
          name: reports
          path: output/
```

## 구현 아키텍처

### 파일 구조

```
pyhub_office_automation/
├── batch/
│   ├── __init__.py
│   ├── parser.py           # Script parser
│   ├── executor.py         # Batch executor
│   ├── variables.py        # Variable management
│   └── control_flow.py     # Control flow (@if, @foreach, @try)
└── cli/
    └── batch_commands.py   # CLI commands for batch execution
```

### 핵심 클래스

#### 1. BatchScript

```python
@dataclass
class BatchScript:
    """Parsed batch script representation"""
    path: str
    lines: List[BatchLine]
    variables: Dict[str, str]
    metadata: Dict[str, str]  # Comments, author, version, etc.

@dataclass
class BatchLine:
    """Single line in batch script"""
    line_number: int
    content: str
    command_type: str  # "shell_command", "directive", "comment", "control"
    command: Optional[str] = None
    args: List[str] = field(default_factory=list)
```

#### 2. BatchExecutor

```python
class BatchExecutor:
    """Execute batch scripts with context management"""

    def __init__(self, script: BatchScript, options: ExecutionOptions):
        self.script = script
        self.options = options
        self.context = UnifiedShellContext()  # Reuse unified shell context
        self.variables = script.variables.copy()
        self.log = []

    def execute(self) -> BatchResult:
        """Execute entire script"""
        for line in self.script.lines:
            if self.should_skip(line):
                continue

            result = self.execute_line(line)
            self.log.append(result)

            if result.failed and not self.options.continue_on_error:
                break

        return BatchResult(
            success=all(r.success for r in self.log),
            executed_lines=len(self.log),
            log=self.log
        )

    def execute_line(self, line: BatchLine) -> LineResult:
        """Execute single line"""
        # Variable substitution
        resolved_line = self.resolve_variables(line)

        # Command execution
        if line.command_type == "directive":
            return self.execute_directive(resolved_line)
        elif line.command_type == "shell_command":
            return self.execute_shell_command(resolved_line)
        elif line.command_type == "control":
            return self.execute_control_flow(resolved_line)
```

#### 3. VariableManager

```python
class VariableManager:
    """Manage script variables and environment"""

    def __init__(self, initial_vars: Dict[str, str] = None):
        self.variables = initial_vars or {}
        self.environment = os.environ.copy()

    def set(self, name: str, value: str):
        """Set variable"""
        self.variables[name] = value

    def get(self, name: str) -> str:
        """Get variable with fallback to environment"""
        return self.variables.get(name, self.environment.get(name, ""))

    def resolve(self, text: str) -> str:
        """Resolve all variables in text: ${VAR} or $VAR"""
        import re

        def replacer(match):
            var_name = match.group(1) or match.group(2)
            return self.get(var_name)

        # ${VAR_NAME} or $VAR_NAME
        pattern = r'\$\{([^}]+)\}|\$([A-Za-z_][A-Za-z0-9_]*)'
        return re.sub(pattern, replacer, text)
```

## 지원 디렉티브

### Variable Management

```bash
@set VAR_NAME = "value"           # 변수 설정
@unset VAR_NAME                   # 변수 삭제
@echo "${VAR_NAME}"               # 변수 출력 (디버깅)
@export VAR_NAME = "value"        # 환경 변수로 내보내기
```

### Control Flow

```bash
@if condition                     # 조건문 시작
  commands...
@elif condition                   # 선택적 조건
  commands...
@else                             # 선택적 기본 분기
  commands...
@endif                            # 조건문 종료

@foreach var in list              # 반복문
  commands...
@endforeach

@while condition                  # While 루프
  commands...
@endwhile
```

### Error Handling

```bash
@try                              # Try-catch 블록
  commands...
@catch                            # 에러 발생 시 실행
  commands...
@finally                          # 항상 실행
  commands...
@endtry

@onerror continue                 # 에러 발생 시 계속 진행
@onerror abort                    # 에러 발생 시 즉시 중단 (기본값)
```

### Script Control

```bash
@run other_script.oas             # 다른 스크립트 실행
@include common_vars.oas          # 스크립트 포함 (변수만)
@sleep 1000                       # 1초 대기 (ms)
@exit                             # 스크립트 종료
```

### File Operations

```bash
@exists "file.xlsx"               # 파일 존재 확인 (조건문용)
@mkdir "output"                   # 디렉토리 생성
@copy "src.xlsx" "dest.xlsx"      # 파일 복사
@delete "temp.xlsx"               # 파일 삭제
```

## 에러 처리 전략

### 1. 기본 동작

```python
class ExecutionOptions:
    continue_on_error: bool = False  # 에러 시 중단
    retry_count: int = 0             # 재시도 횟수
    retry_delay: int = 1000          # 재시도 간격 (ms)
    log_file: Optional[str] = None   # 로그 파일
    verbose: bool = False            # 상세 로그
    dry_run: bool = False            # 실제 실행하지 않음
```

### 2. 로깅

```python
@dataclass
class LineResult:
    """Single line execution result"""
    line_number: int
    command: str
    success: bool
    output: str
    error: Optional[str] = None
    duration_ms: int = 0

@dataclass
class BatchResult:
    """Overall batch execution result"""
    success: bool
    executed_lines: int
    skipped_lines: int
    failed_lines: int
    total_duration_ms: int
    log: List[LineResult]
    start_time: datetime
    end_time: datetime
```

## 실전 예제

### 예제 1: 주간 판매 보고서 자동화

```bash
# weekly_sales_report.oas
# 매주 월요일 자동 실행

@set WEEK_START = "2024-01-01"
@set WEEK_END = "2024-01-07"
@set REPORT_FILE = "weekly_report_${WEEK_START}.pptx"

# 데이터 처리
use excel "sales_data.xlsx"
use sheet "Transactions"
range-read --range "A1:G1000" --output-file "weekly_data.csv"

# 집계
sheet-add --name "Weekly_Summary"
use sheet "Weekly_Summary"
range-write --range "A1" --data '[["Week","Total Sales","Orders"],["${WEEK_START}",50000,120]]'

# 차트 생성
chart-add --data-range "A1:C2" --chart-type "Column" --title "Weekly Performance"

# PowerPoint 보고서
use ppt "${REPORT_FILE}"
slide-add --layout 1
use slide 2
content-add-text --text "Weekly Sales Report - ${WEEK_START}" --left 100 --top 50
content-add-excel-chart --excel-file "sales_data.xlsx" --sheet "Weekly_Summary" --chart-name "Chart1"
```

### 예제 2: 다중 파일 처리

```bash
# process_quarterly_files.oas
# Q1, Q2, Q3, Q4 데이터 통합

@set QUARTERS = ["Q1", "Q2", "Q3", "Q4"]
@set OUTPUT_FILE = "annual_summary.xlsx"

# 출력 워크북 생성
use excel "${OUTPUT_FILE}"
sheet-add --name "Summary"

@foreach quarter in ${QUARTERS}
  @echo "Processing ${quarter}..."

  @if exists("${quarter}_sales.xlsx")
    # 분기별 데이터 읽기
    use excel "${quarter}_sales.xlsx"
    use sheet "Data"
    table-read --output-file "${quarter}_data.csv"

    # 요약 시트에 추가
    use excel "${OUTPUT_FILE}"
    use sheet "Summary"
    range-write --range "A${__LOOP_INDEX__}" --data-file "${quarter}_data.csv"
  @else
    @echo "Warning: ${quarter}_sales.xlsx not found, skipping..."
  @endif
@endforeach

# 최종 차트 생성
chart-add --data-range "A1:D10" --chart-type "Line" --title "Quarterly Trends"
```

### 예제 3: 에러 처리 및 재시도

```bash
# robust_workflow.oas
# 에러에 강한 워크플로우

@onerror continue  # 에러 발생해도 계속 진행

@try
  use excel "data.xlsx"
  use sheet "Data"
@catch
  @echo "Failed to open data.xlsx, using backup..."
  use excel "backup_data.xlsx"
  use sheet "Data"
@endtry

# 여러 차트 생성 시도
@foreach chart_type in ["Column", "Line", "Pie"]
  @try
    chart-add --data-range "A1:B10" --chart-type "${chart_type}" --title "${chart_type} Chart"
  @catch
    @echo "Failed to create ${chart_type} chart, continuing..."
  @endtry
@endforeach
```

## CLI 명령어

```bash
# 스크립트 실행
oa batch run <script.oas>              # 기본 실행
oa batch run <script.oas> --dry-run    # 실행 시뮬레이션
oa batch run <script.oas> --verbose    # 상세 로그
oa batch run <script.oas> --set VAR=value  # 변수 오버라이드

# 스크립트 검증
oa batch validate <script.oas>         # 문법 검사
oa batch format <script.oas>           # 자동 포맷팅

# 스크립트 관리
oa batch list                          # 스크립트 목록
oa batch info <script.oas>             # 스크립트 정보 조회
```

## 장점

1. **재현성**: 작업 과정을 정확히 재현 가능
2. **자동화**: 반복 작업을 스크립트로 자동화
3. **문서화**: 스크립트 자체가 작업 문서
4. **CI/CD 통합**: 자동화 파이프라인에 쉽게 통합
5. **에러 처리**: 체계적인 에러 처리 및 재시도

## 제한사항

1. **대화형 명령**: `input()` 등 대화형 명령은 지원 안 됨
2. **복잡한 로직**: 매우 복잡한 로직은 Python 스크립트 권장
3. **플랫폼 의존성**: Windows 전용 기능은 macOS에서 실행 불가

---

**Design Version**: 1.0
**Author**: Claude Code
**Date**: 2025-10-01
