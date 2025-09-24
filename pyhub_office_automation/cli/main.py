"""
pyhub-office-automation Typer 기반 CLI 명령어
PyInstaller 호환성을 위한 정적 명령어 등록
"""

import json
import os
import sys
from typing import Optional

# Windows 환경에서 UTF-8 인코딩 강제 설정
if sys.platform == "win32":
    # 환경 변수 설정
    os.environ.setdefault("PYTHONIOENCODING", "utf-8")
    os.environ.setdefault("PYTHONUTF8", "1")

    # stdout/stderr 인코딩 설정
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8")
            sys.stderr.reconfigure(encoding="utf-8")
        except Exception:
            pass  # 설정 실패해도 계속 진행

import typer
from rich.console import Console
from rich.table import Table

from pyhub_office_automation.cli.ai_setup import ai_setup_app

# Chart 명령어 import
from pyhub_office_automation.excel.chart_add import chart_add
from pyhub_office_automation.excel.chart_configure import chart_configure
from pyhub_office_automation.excel.chart_delete import chart_delete
from pyhub_office_automation.excel.chart_export import chart_export
from pyhub_office_automation.excel.chart_list import chart_list
from pyhub_office_automation.excel.chart_pivot_create import chart_pivot_create
from pyhub_office_automation.excel.chart_position import chart_position

# Data 명령어 import (Issue #39)
from pyhub_office_automation.excel.data_analyze import data_analyze
from pyhub_office_automation.excel.data_transform import data_transform
from pyhub_office_automation.excel.metadata_generate import metadata_generate

# Pivot 명령어 import
from pyhub_office_automation.excel.pivot_configure import pivot_configure
from pyhub_office_automation.excel.pivot_create import pivot_create
from pyhub_office_automation.excel.pivot_delete import pivot_delete
from pyhub_office_automation.excel.pivot_list import pivot_list
from pyhub_office_automation.excel.pivot_refresh import pivot_refresh

# Excel 명령어 import
from pyhub_office_automation.excel.range_convert import range_convert
from pyhub_office_automation.excel.range_read import range_read
from pyhub_office_automation.excel.range_write import range_write

# Shape 명령어 import
from pyhub_office_automation.excel.shape_add import shape_add
from pyhub_office_automation.excel.shape_delete import shape_delete
from pyhub_office_automation.excel.shape_format import shape_format
from pyhub_office_automation.excel.shape_group import shape_group
from pyhub_office_automation.excel.shape_list import shape_list
from pyhub_office_automation.excel.sheet_activate import sheet_activate
from pyhub_office_automation.excel.sheet_add import sheet_add
from pyhub_office_automation.excel.sheet_delete import sheet_delete
from pyhub_office_automation.excel.sheet_rename import sheet_rename

# Slicer 명령어 import
from pyhub_office_automation.excel.slicer_add import slicer_add
from pyhub_office_automation.excel.slicer_connect import slicer_connect
from pyhub_office_automation.excel.slicer_list import slicer_list
from pyhub_office_automation.excel.slicer_position import slicer_position
from pyhub_office_automation.excel.table_analyze import table_analyze
from pyhub_office_automation.excel.table_create import table_create
from pyhub_office_automation.excel.table_list import table_list
from pyhub_office_automation.excel.table_read import table_read
from pyhub_office_automation.excel.table_sort import table_sort
from pyhub_office_automation.excel.table_sort_clear import table_sort_clear
from pyhub_office_automation.excel.table_sort_info import table_sort_info
from pyhub_office_automation.excel.table_write import table_write
from pyhub_office_automation.excel.textbox_add import textbox_add
from pyhub_office_automation.excel.workbook_create import workbook_create
from pyhub_office_automation.excel.workbook_info import workbook_info
from pyhub_office_automation.excel.workbook_list import workbook_list
from pyhub_office_automation.excel.workbook_open import workbook_open

# HWP 명령어 import
from pyhub_office_automation.hwp.hwp_export import hwp_export
from pyhub_office_automation.mcp.cli import mcp_app
from pyhub_office_automation.utils.resource_loader import load_llm_guide, load_welcome_message
from pyhub_office_automation.version import get_version, get_version_info

# Typer 앱 생성
app = typer.Typer(help="pyhub-office-automation: AI 에이전트를 위한 Office 자동화 도구")


def version_callback(value: bool):
    """--version 콜백 함수"""
    if value:
        version_info = get_version_info()
        typer.echo(f"pyhub-office-automation version {version_info['version']}")
        raise typer.Exit()


# 글로벌 --version 옵션 추가 및 기본 메시지 표시
@app.callback(invoke_without_command=True)
def main_callback(
    ctx: typer.Context,
    version: bool = typer.Option(False, "--version", "-v", callback=version_callback, help="버전 정보 출력"),
):
    """
    pyhub-office-automation: AI 에이전트를 위한 Office 자동화 도구
    """
    # 서브커맨드가 없고 버전 옵션도 아닌 경우 welcome 메시지 표시
    if ctx.invoked_subcommand is None:
        show_welcome_message()


def show_welcome_message():
    """Welcome 메시지를 표시합니다."""
    welcome_content = load_welcome_message()
    console.print(welcome_content)

    # LLM 가이드 안내 추가
    console.print("\n💡 [bold cyan]AI 에이전트 사용 시 상세 지침을 보려면:[/bold cyan]")
    console.print("   oa llm-guide")


# version 명령어 추가
@app.command()
def version():
    """버전 정보 출력"""
    version_info = get_version_info()
    typer.echo(f"pyhub-office-automation version {version_info['version']}")


@app.command()
def welcome(output_format: str = typer.Option("text", "--format", help="출력 형식 선택 (text, json)")):
    """환영 메시지 및 시작 가이드 출력"""
    welcome_content = load_welcome_message()

    if output_format == "json":
        # JSON 형식으로 출력
        welcome_data = {
            "message_type": "welcome",
            "content": welcome_content,
            "package_version": get_version(),
            "available_commands": {
                "info": "패키지 정보 및 설치 상태",
                "excel": "Excel 자동화 명령어들",
                "hwp": "HWP 자동화 명령어들 (Windows 전용)",
                "install-guide": "설치 가이드",
                "llm-guide": "AI 에이전트 사용 지침",
            },
        }
        try:
            json_output = json.dumps(welcome_data, ensure_ascii=False, indent=2)
            typer.echo(json_output)
        except UnicodeEncodeError:
            json_output = json.dumps(welcome_data, ensure_ascii=True, indent=2)
            typer.echo(json_output)
    else:
        console.print(welcome_content)


@app.command()
def llm_guide(output_format: str = typer.Option("text", "--format", help="출력 형식 선택 (text, json, markdown)")):
    """LLM/AI 에이전트를 위한 상세 사용 지침"""
    guide_content = load_llm_guide()

    if output_format == "json":
        # JSON 형식으로 출력 (AI 에이전트가 파싱하기 쉽도록)
        guide_data = {
            "guide_type": "llm_usage",
            "content": guide_content,
            "package_version": get_version(),
            "target_audience": "LLM, AI Agent, Chatbot",
            "key_principles": [
                "명령어 발견 (Command Discovery)",
                "컨텍스트 인식 (Context Awareness)",
                "에러 방지 워크플로우",
                "효율적인 연결 방법 활용",
            ],
            "essential_commands": {
                "discovery": ["oa info", "oa excel list --format json", "oa hwp list --format json"],
                "context": ["oa excel workbook-list", "oa excel workbook-info --include-sheets"],
                "workflow": ["연속 작업시 활성 워크북 자동 사용 또는 --workbook-name 사용"],
            },
            "connection_methods": [
                "--file-path: 파일 경로로 연결",
                "옵션 없음: 활성 워크북 자동 사용 (기본값)",
                "--workbook-name: 워크북 이름으로 연결",
            ],
        }
        try:
            json_output = json.dumps(guide_data, ensure_ascii=False, indent=2)
            typer.echo(json_output)
        except UnicodeEncodeError:
            json_output = json.dumps(guide_data, ensure_ascii=True, indent=2)
            typer.echo(json_output)
    elif output_format == "markdown":
        # 원본 마크다운 출력
        typer.echo(guide_content)
    else:
        # 콘솔에 포맷팅된 출력
        console.print(guide_content)


excel_app = typer.Typer(help="Excel 자동화 명령어들", no_args_is_help=True)
hwp_app = typer.Typer(help="HWP 자동화 명령어들 (Windows 전용)", no_args_is_help=True)

# Rich 콘솔 - UTF-8 인코딩 안전성 확보
try:
    # Windows 환경에서 UTF-8 출력 보장
    console = Console(force_terminal=True, force_jupyter=False, legacy_windows=False, width=None)  # 자동 감지
except Exception:
    # fallback to basic console
    console = Console(legacy_windows=True)

# Excel 명령어 등록 (단계적 테스트)
# Range Commands
excel_app.command("range-read")(range_read)
excel_app.command("range-write")(range_write)
excel_app.command("range-convert")(range_convert)

# Data Commands (Issue #39)
excel_app.command("data-analyze")(data_analyze)
excel_app.command("data-transform")(data_transform)

# Workbook Commands
excel_app.command("workbook-list")(workbook_list)
excel_app.command("workbook-open")(workbook_open)
excel_app.command("workbook-create")(workbook_create)
excel_app.command("workbook-info")(workbook_info)
excel_app.command("metadata-generate")(metadata_generate)

# Sheet Commands
excel_app.command("sheet-activate")(sheet_activate)
excel_app.command("sheet-add")(sheet_add)
excel_app.command("sheet-delete")(sheet_delete)
excel_app.command("sheet-rename")(sheet_rename)

# Table Commands
excel_app.command("table-create")(table_create)
excel_app.command("table-list")(table_list)
excel_app.command("table-read")(table_read)
excel_app.command("table-sort")(table_sort)
excel_app.command("table-sort-clear")(table_sort_clear)
excel_app.command("table-sort-info")(table_sort_info)
excel_app.command("table-write")(table_write)
excel_app.command("table-analyze")(table_analyze)

# Chart Commands
excel_app.command("chart-add")(chart_add)
excel_app.command("chart-configure")(chart_configure)
excel_app.command("chart-delete")(chart_delete)
excel_app.command("chart-export")(chart_export)
excel_app.command("chart-list")(chart_list)
excel_app.command("chart-pivot-create")(chart_pivot_create)
excel_app.command("chart-position")(chart_position)

# Pivot Commands
excel_app.command("pivot-configure")(pivot_configure)
excel_app.command("pivot-create")(pivot_create)
excel_app.command("pivot-delete")(pivot_delete)
excel_app.command("pivot-list")(pivot_list)
excel_app.command("pivot-refresh")(pivot_refresh)

# Shape Commands (이제 Typer로 전환 완료)
excel_app.command("shape-add")(shape_add)
excel_app.command("shape-delete")(shape_delete)
excel_app.command("shape-format")(shape_format)
excel_app.command("shape-group")(shape_group)
excel_app.command("shape-list")(shape_list)
excel_app.command("textbox-add")(textbox_add)

# Slicer Commands (이제 Typer로 전환 완료)
excel_app.command("slicer-add")(slicer_add)
excel_app.command("slicer-connect")(slicer_connect)
excel_app.command("slicer-list")(slicer_list)
excel_app.command("slicer-position")(slicer_position)


# Excel list command
@excel_app.command("list")
def excel_list_temp(
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택"),
):
    """Excel 자동화 명령어 목록 출력"""
    commands = [
        # Workbook Commands
        {"name": "workbook-list", "description": "열린 Excel 워크북 목록 조회", "category": "workbook"},
        {"name": "workbook-open", "description": "Excel 워크북 열기", "category": "workbook"},
        {"name": "workbook-create", "description": "새 Excel 워크북 생성", "category": "workbook"},
        {"name": "workbook-info", "description": "워크북 정보 조회", "category": "workbook"},
        {"name": "metadata-generate", "description": "워크북 전체 Excel Table 메타데이터 자동 생성", "category": "workbook"},
        # Sheet Commands
        {"name": "sheet-activate", "description": "시트 활성화", "category": "sheet"},
        {"name": "sheet-add", "description": "새 시트 추가", "category": "sheet"},
        {"name": "sheet-delete", "description": "시트 삭제", "category": "sheet"},
        {"name": "sheet-rename", "description": "시트 이름 변경", "category": "sheet"},
        # Range Commands
        {"name": "range-read", "description": "셀 범위 데이터 읽기", "category": "range"},
        {"name": "range-write", "description": "셀 범위에 데이터 쓰기", "category": "range"},
        {"name": "range-convert", "description": "셀 범위 데이터 형식 변환 (문자열 → 숫자)", "category": "range"},
        # Data Commands (Issue #39)
        {"name": "data-analyze", "description": "피벗테이블용 데이터 구조 분석", "category": "data"},
        {"name": "data-transform", "description": "피벗테이블용 형식으로 데이터 변환", "category": "data"},
        # Table Commands
        {"name": "table-create", "description": "기존 범위를 Excel Table로 변환", "category": "table"},
        {"name": "table-list", "description": "워크북의 Excel Table 목록 조회", "category": "table"},
        {"name": "table-read", "description": "테이블 데이터를 DataFrame으로 읽기", "category": "table"},
        {"name": "table-sort", "description": "테이블 정렬 적용", "category": "table"},
        {"name": "table-sort-clear", "description": "테이블 정렬 해제", "category": "table"},
        {"name": "table-sort-info", "description": "테이블 정렬 상태 확인", "category": "table"},
        {"name": "table-write", "description": "DataFrame을 Excel 테이블로 쓰기 (선택적 Table 생성)", "category": "table"},
        {"name": "table-analyze", "description": "Excel Table 분석 및 메타데이터 자동 생성", "category": "table"},
        # Chart Commands
        {"name": "chart-add", "description": "차트 추가", "category": "chart"},
        {"name": "chart-configure", "description": "차트 설정", "category": "chart"},
        {"name": "chart-delete", "description": "차트 삭제", "category": "chart"},
        {"name": "chart-export", "description": "차트 내보내기", "category": "chart"},
        {"name": "chart-list", "description": "차트 목록 조회", "category": "chart"},
        {"name": "chart-pivot-create", "description": "피벗 차트 생성", "category": "chart"},
        {"name": "chart-position", "description": "차트 위치 설정", "category": "chart"},
        # Pivot Commands
        {"name": "pivot-configure", "description": "피벗테이블 설정", "category": "pivot"},
        {"name": "pivot-create", "description": "피벗테이블 생성", "category": "pivot"},
        {"name": "pivot-delete", "description": "피벗테이블 삭제", "category": "pivot"},
        {"name": "pivot-list", "description": "피벗테이블 목록 조회", "category": "pivot"},
        {"name": "pivot-refresh", "description": "피벗테이블 새로고침", "category": "pivot"},
        # Shape Commands
        {"name": "shape-add", "description": "도형 추가", "category": "shape"},
        {"name": "shape-delete", "description": "도형 삭제", "category": "shape"},
        {"name": "shape-format", "description": "도형 서식 설정", "category": "shape"},
        {"name": "shape-group", "description": "도형 그룹화", "category": "shape"},
        {"name": "shape-list", "description": "도형 목록 조회", "category": "shape"},
        {"name": "textbox-add", "description": "텍스트 상자 추가", "category": "shape"},
        # Slicer Commands
        {"name": "slicer-add", "description": "슬라이서 추가", "category": "slicer"},
        {"name": "slicer-connect", "description": "슬라이서 연결", "category": "slicer"},
        {"name": "slicer-list", "description": "슬라이서 목록 조회", "category": "slicer"},
        {"name": "slicer-position", "description": "슬라이서 위치 설정", "category": "slicer"},
    ]

    excel_data = {
        "category": "excel",
        "description": "Excel 자동화 명령어들 (xlwings 기반)",
        "platform_requirement": "Windows (전체 기능) / macOS (제한적)",
        "commands": commands,
        "total_commands": len(commands),
        "package_version": get_version(),
    }

    if output_format == "json":
        try:
            json_output = json.dumps(excel_data, ensure_ascii=False, indent=2)
            typer.echo(json_output)
        except UnicodeEncodeError:
            json_output = json.dumps(excel_data, ensure_ascii=True, indent=2)
            typer.echo(json_output)
    else:
        console.print("=== Excel 자동화 명령어 목록 ===", style="bold green")
        console.print(f"Platform: {excel_data['platform_requirement']}")
        console.print(f"Total: {excel_data['total_commands']} commands")
        console.print()

        categories = {}
        for cmd in commands:
            category = cmd["category"]
            if category not in categories:
                categories[category] = []
            categories[category].append(cmd)

        for category, cmds in categories.items():
            console.print(f"[bold blue]{category.upper()} Commands:[/bold blue]")
            for cmd in cmds:
                console.print(f"  • oa excel {cmd['name']}")
                console.print(f"    {cmd['description']}")
            console.print()

        console.print("📚 [bold yellow]더 자세한 사용 지침은 다음 명령어를 참고하세요:[/bold yellow]", style="bright_yellow")
        console.print("   [bold cyan]oa llm-guide[/bold cyan] - AI 에이전트를 위한 상세 가이드")
        console.print("   [bold cyan]oa excel <command> --help[/bold cyan] - 특정 명령어 도움말")
        console.print()


# 서브 앱을 메인 앱에 등록
app.add_typer(excel_app, name="excel")
app.add_typer(hwp_app, name="hwp")
app.add_typer(ai_setup_app, name="ai-setup")
app.add_typer(mcp_app, name="mcp")


@app.command()
def info(
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택"),
):
    """패키지 정보 및 설치 상태 출력"""
    try:
        version_info = get_version_info()
        dependencies = check_dependencies()

        info_data = {
            "package": "pyhub-office-automation",
            "version": version_info,
            "platform": sys.platform,
            "python_version": sys.version,
            "dependencies": dependencies,
            "status": "installed",
        }

        if output_format == "json":
            try:
                json_output = json.dumps(info_data, ensure_ascii=False, indent=2)
                typer.echo(json_output)
            except UnicodeEncodeError:
                json_output = json.dumps(info_data, ensure_ascii=True, indent=2)
                typer.echo(json_output)
        else:
            console.print(f"Package: {info_data['package']}", style="bold green")
            console.print(f"Version: {info_data['version']['version']}")
            console.print(f"Platform: {info_data['platform']}")
            console.print(f"Python: {info_data['python_version']}")
            console.print("Dependencies:", style="bold")

            table = Table()
            table.add_column("Package", style="cyan")
            table.add_column("Status", style="green")
            table.add_column("Version")

            for dep, status in info_data["dependencies"].items():
                status_mark = "✓" if status["available"] else "✗"
                version = status["version"] or "Not installed"
                table.add_row(dep, status_mark, version)

            console.print(table)

    except Exception as e:
        error_data = {"error": str(e), "command": "info", "version": get_version()}
        typer.echo(json.dumps(error_data, ensure_ascii=False, indent=2), err=True)
        raise typer.Exit(1)


@app.command()
def install_guide(
    output_format: str = typer.Option("text", "--format", help="출력 형식 선택"),
):
    """설치 가이드 출력"""
    guide_steps = [
        {
            "step": 1,
            "title": "Python 설치",
            "description": "Python 3.13 이상을 설치하세요",
            "url": "https://www.python.org/downloads/",
            "command": None,
        },
        {
            "step": 2,
            "title": "패키지 설치",
            "description": "pip를 사용하여 pyhub-office-automation을 설치하세요",
            "command": "pip install pyhub-office-automation",
        },
        {"step": 3, "title": "설치 확인", "description": "oa 명령어가 정상 동작하는지 확인하세요", "command": "oa info"},
        {
            "step": 4,
            "title": "Excel 사용 시 (선택사항)",
            "description": "Microsoft Excel이 설치되어 있어야 합니다",
            "note": "xlwings는 Excel이 설치된 환경에서 동작합니다",
        },
        {
            "step": 5,
            "title": "HWP 사용 시 (선택사항, Windows 전용)",
            "description": "한글(HWP) 프로그램이 설치되어 있어야 합니다",
            "note": "pyhwpx는 Windows COM을 통해 HWP와 연동됩니다",
        },
    ]

    guide_data = {
        "title": "pyhub-office-automation 설치 가이드",
        "version": get_version(),
        "platform_requirement": "Windows 10/11 (추천)",
        "python_requirement": "Python 3.13+",
        "steps": guide_steps,
    }

    if output_format == "json":
        try:
            # JSON 출력 시 한글 인코딩 문제 해결
            json_output = json.dumps(guide_data, ensure_ascii=False, indent=2)
            typer.echo(json_output)
        except UnicodeEncodeError:
            # Windows 콘솔 인코딩 문제 시 ensure_ascii=True로 폴백
            json_output = json.dumps(guide_data, ensure_ascii=True, indent=2)
            typer.echo(json_output)
    else:
        console.print(f"=== {guide_data['title']} ===", style="bold blue")
        console.print(f"Version: {guide_data['version']}")
        console.print(f"Platform: {guide_data['platform_requirement']}")
        console.print(f"Python: {guide_data['python_requirement']}")
        console.print()

        for step in guide_steps:
            console.print(f"Step {step['step']}: {step['title']}", style="bold yellow")
            console.print(f"  {step['description']}")
            if step.get("command"):
                console.print(f"  Command: [green]{step['command']}[/green]")
            if step.get("url"):
                console.print(f"  URL: [blue]{step['url']}[/blue]")
            if step.get("note"):
                console.print(f"  Note: [dim]{step['note']}[/dim]")
            console.print()


@hwp_app.command("list")
def hwp_list(
    output_format: str = typer.Option("json", "--format", help="출력 형식 선택"),
):
    """HWP 자동화 명령어 목록 출력"""
    commands = [
        {"name": "export", "description": "HWP 파일을 HTML로 변환", "version": "1.0.0", "status": "available"},
        {"name": "open", "description": "HWP 파일 열기", "version": "1.0.0", "status": "planned"},
        {"name": "save", "description": "HWP 파일 저장", "version": "1.0.0", "status": "planned"},
    ]

    hwp_data = {
        "category": "hwp",
        "description": "HWP 자동화 명령어들 (pyhwpx 기반, Windows 전용)",
        "platform_requirement": "Windows + HWP 프로그램 설치 필요",
        "commands": commands,
        "total_commands": len(commands),
        "package_version": get_version(),
    }

    if output_format == "json":
        try:
            json_output = json.dumps(hwp_data, ensure_ascii=False, indent=2)
            typer.echo(json_output)
        except UnicodeEncodeError:
            json_output = json.dumps(hwp_data, ensure_ascii=True, indent=2)
            typer.echo(json_output)
    else:
        console.print("=== HWP 자동화 명령어 목록 ===", style="bold blue")
        console.print(f"Platform: {hwp_data['platform_requirement']}")
        console.print(f"Total: {hwp_data['total_commands']} commands")
        console.print()

        for cmd in commands:
            status_mark = "✓" if cmd["status"] == "available" else "○"
            console.print(f"  {status_mark} oa hwp {cmd['name']}")
            console.print(f"     {cmd['description']} (v{cmd['version']})")


@hwp_app.command("export")
def hwp_export_command(
    file_path: str = typer.Option(..., "--file-path", help="변환할 HWP 파일의 절대 경로"),
    format_type: str = typer.Option("html", "--format", help="출력 형식 (현재 html만 지원)"),
    output_file: Optional[str] = typer.Option(None, "--output-file", help="HTML 저장 경로 (선택, 미지정시 표준출력)"),
    encoding: str = typer.Option("utf-8", "--encoding", help="출력 인코딩 (기본값: utf-8)"),
    include_css: bool = typer.Option(
        False, "--include-css/--no-include-css", help="CSS 스타일 포함 여부 (기본값: False, 모든 CSS 제거)"
    ),
    include_images: bool = typer.Option(
        False, "--include-images/--no-include-images", help="이미지 포함 여부 (기본값: False, Base64 인코딩으로 포함)"
    ),
    temp_cleanup: bool = typer.Option(True, "--temp-cleanup/--no-temp-cleanup", help="임시 파일 자동 정리 (기본값: True)"),
    output_format: str = typer.Option("json", "--output-format", help="응답 출력 형식 (json)"),
):
    """HWP 파일을 HTML 형식으로 변환"""
    hwp_export(
        file_path=file_path,
        format_type=format_type,
        output_file=output_file,
        encoding=encoding,
        include_css=include_css,
        include_images=include_images,
        temp_cleanup=temp_cleanup,
        output_format=output_format,
    )


@app.command()
def get_help(
    category: str = typer.Argument(..., help="명령어 카테고리 (excel, hwp)"),
    command_name: str = typer.Argument(..., help="명령어 이름"),
    output_format: str = typer.Option("text", "--format", help="출력 형식 선택"),
):
    """특정 명령어의 도움말 조회"""
    help_data = {
        "category": category,
        "command": command_name,
        "help": f"oa {category} {command_name} 명령어 도움말 (구현 예정)",
        "usage": f"oa {category} {command_name} [OPTIONS]",
        "status": "planned",
        "version": get_version(),
    }

    if output_format == "json":
        try:
            json_output = json.dumps(help_data, ensure_ascii=False, indent=2)
            typer.echo(json_output)
        except UnicodeEncodeError:
            json_output = json.dumps(help_data, ensure_ascii=True, indent=2)
            typer.echo(json_output)
    else:
        console.print(f"Command: oa {category} {command_name}", style="bold")
        console.print(f"Usage: {help_data['usage']}")
        console.print(f"Status: {help_data['status']}")
        console.print()
        console.print(help_data["help"])


def check_dependencies():
    """의존성 패키지 설치 상태 확인"""
    dependencies = {}

    # xlwings 확인
    try:
        import xlwings

        dependencies["xlwings"] = {"available": True, "version": xlwings.__version__}
    except ImportError:
        dependencies["xlwings"] = {"available": False, "version": None}

    # pyhwpx 확인 (Windows 전용)
    try:
        import pyhwpx

        dependencies["pyhwpx"] = {"available": True, "version": getattr(pyhwpx, "__version__", "unknown")}
    except ImportError:
        dependencies["pyhwpx"] = {"available": False, "version": None}

    # pandas 확인
    try:
        import pandas

        dependencies["pandas"] = {"available": True, "version": pandas.__version__}
    except ImportError:
        dependencies["pandas"] = {"available": False, "version": None}

    return dependencies


def main():
    """메인 엔트리포인트"""
    app()


if __name__ == "__main__":
    main()
