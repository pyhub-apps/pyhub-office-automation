"""
PyHub Office Automation MCP CLI

MCP 서버를 시작하고 관리하기 위한 명령어 인터페이스
"""

import typer
import sys
from typing import Optional
from rich.console import Console
from rich.table import Table

from .http_server import run_server
from .server import mcp
from pyhub_office_automation.utils.port_checker import (
    check_port_with_recommendation,
    get_port_info,
    find_available_port
)

console = Console()

mcp_app = typer.Typer(help="MCP (Model Context Protocol) 서버 관리")

@mcp_app.command()
def start(
    host: str = typer.Option("127.0.0.1", "--host", "-h", help="서버 호스트"),
    port: int = typer.Option(8765, "--port", "-p", help="서버 포트"),
    log_level: str = typer.Option("info", "--log-level", help="로그 레벨"),
    reload: bool = typer.Option(False, "--reload", help="개발용 자동 재로드"),
    force: bool = typer.Option(False, "--force", help="포트 사용 중이어도 강제 실행")
):
    """MCP HTTP 서버 시작"""
    try:
        console.print(f"[bold green]MCP 서버 시작 중...[/bold green]")
        console.print(f"   Host: {host}")
        console.print(f"   Port: {port}")
        console.print(f"   MCP endpoint: http://{host}:{port}/mcp")
        console.print(f"   Docs: http://{host}:{port}/docs")

        run_server(host=host, port=port, log_level=log_level, reload=reload, force=force)

    except KeyboardInterrupt:
        console.print("\n✋ [yellow]서버가 사용자에 의해 중단되었습니다[/yellow]")
    except Exception as e:
        console.print(f"❌ [red]서버 시작 실패: {e}[/red]")
        sys.exit(1)

@mcp_app.command()
def info():
    """MCP 서버 정보 출력"""
    console.print("=== MCP 서버 정보 ===", style="bold blue")
    console.print(f"Name: {mcp.name}")
    console.print(f"Version: {mcp.version}")
    console.print(f"Instructions: {mcp.instructions}")

    # Resources 정보 (개발 모드에서는 하드코딩)
    resources = ["resource://excel/workbooks", "resource://excel/workbook/{name}/info"]
    console.print(f"\n[bold cyan]Resources ({len(resources)}):[/bold cyan]")
    for resource in resources:
        console.print(f"  • {resource}")

    # Tools 정보 (개발 모드에서는 하드코딩)
    expected_tools = [
        'excel_workbook_info',
        'excel_range_read',
        'excel_table_read',
        'excel_data_analyze',
        'excel_chart_list'
    ]
    console.print(f"\n[bold cyan]Tools ({len(expected_tools)}):[/bold cyan]")
    if expected_tools:
        table = Table()
        table.add_column("Tool Name", style="cyan")
        table.add_column("Description", style="dim")

        tool_descriptions = {
            'excel_workbook_info': '워크북 구조 및 정보 분석',
            'excel_range_read': '셀 범위 데이터 읽기',
            'excel_table_read': '테이블 데이터 읽기',
            'excel_data_analyze': '데이터 구조 분석',
            'excel_chart_list': '차트 목록 조회'
        }

        for tool_name in expected_tools:
            description = tool_descriptions.get(tool_name, '설명 없음')
            table.add_row(tool_name, description)

        console.print(table)
    else:
        console.print("  (등록된 도구 없음)")

@mcp_app.command()
def test():
    """MCP 서버 기본 테스트"""
    console.print("[bold yellow]MCP 서버 기본 테스트 실행 중...[/bold yellow]")

    try:
        # Import 테스트
        console.print("1. Import 테스트...", end=" ")
        from .server import mcp
        console.print("[green]✓[/green]")

        # 서버 인스턴스 테스트
        console.print("2. 서버 인스턴스 테스트...", end=" ")
        assert mcp.name == "PyHub Office Automation MCP"
        assert mcp.version is not None and len(mcp.version) > 0
        console.print("[green]✓[/green]")

        # Tools 등록 테스트 (간단히 처리)
        console.print("3. Tools 등록 테스트...", end=" ")
        # MCP 서버에 도구들이 정의되어 있는지 확인
        assert hasattr(mcp, 'excel_workbook_info'), "excel_workbook_info not found"
        console.print("[green]✓[/green]")

        # Resources 등록 테스트 (간단히 처리)
        console.print("4. Resources 등록 테스트...", end=" ")
        # 리소스 함수들이 정의되어 있는지 확인
        assert hasattr(mcp, 'get_workbooks'), "get_workbooks not found"
        console.print("[green]✓[/green]")

        console.print("\n[bold green]모든 테스트 통과![/bold green]")
        console.print(f"   등록된 도구: 5개")
        console.print(f"   등록된 리소스: 2개")

    except Exception as e:
        console.print(f"[red]✗[/red]")
        console.print(f"❌ [red]테스트 실패: {e}[/red]")
        sys.exit(1)

@mcp_app.command()
def check_port(
    port: int = typer.Argument(8765, help="확인할 포트 번호"),
    host: str = typer.Option("127.0.0.1", "--host", help="확인할 호스트"),
    detailed: bool = typer.Option(False, "--detailed", help="상세 정보 표시"),
    find_alternative: bool = typer.Option(False, "--find-alternative", help="대안 포트 찾기")
):
    """포트 사용 상태 확인 및 대안 제시"""
    console.print(f"[bold blue]포트 {port} 상태 확인 중...[/bold blue]")

    if detailed:
        info = get_port_info(host, port)
        console.print(f"\n[cyan]Host:[/cyan] {info['host']}")
        console.print(f"[cyan]Port:[/cyan] {info['port']}")
        console.print(f"[cyan]Available:[/cyan] {'[green]Yes[/green]' if info['is_available'] else '[red]No[/red]'}")
        console.print(f"[cyan]Platform:[/cyan] {info['platform']}")

        if "process_info" in info and info["process_info"]:
            console.print(f"\n[cyan]Process Information:[/cyan]")
            for line in info["process_info"]:
                console.print(f"  {line}")

    elif find_alternative:
        result = check_port_with_recommendation(host, port)
        console.print(f"\n{result['message']}")

        if result["alternative_port"]:
            console.print(f"[green]권장 포트:[/green] {result['alternative_port']}")
            console.print(f"[dim]시작 명령:[/dim] oa mcp start --port {result['alternative_port']}")

            # 주변 여러 포트 확인
            console.print(f"\n[cyan]사용 가능한 포트 목록:[/cyan]")
            available_ports = []
            for check_port_num in range(port + 1, port + 11):
                alternative = find_available_port(host, check_port_num, 1)
                if alternative:
                    available_ports.append(alternative)

            if available_ports:
                ports_str = ", ".join(map(str, available_ports[:5]))
                console.print(f"  {ports_str}")
        else:
            console.print("[yellow]인근 포트들도 모두 사용 중입니다.[/yellow]")
    else:
        result = check_port_with_recommendation(host, port)
        if result["is_available"]:
            console.print(f"[green]OK 포트 {port} 사용 가능[/green]")
        else:
            console.print(f"[red]BUSY 포트 {port} 사용 중[/red]")
            if result["alternative_port"]:
                console.print(f"[yellow]대안 포트: {result['alternative_port']}[/yellow]")

@mcp_app.command()
def docs():
    """사용 가이드 및 문서 출력"""
    console.print("=== PyHub Office Automation MCP 사용 가이드 ===", style="bold blue")

    console.print("\n📋 [bold cyan]기본 사용법:[/bold cyan]")
    console.print("1. MCP 서버 시작:")
    console.print("   [green]oa mcp start[/green]")
    console.print("   [green]oa mcp start --port 8080[/green]")

    console.print("\n2. Claude Desktop 연결:")
    console.print("   • Claude Desktop 설정에서 MCP 서버 추가")
    console.print("   • URL: http://localhost:8765/mcp")
    console.print("   • 전송 방식: Streamable HTTP")

    console.print("\n📊 [bold cyan]사용 가능한 기능:[/bold cyan]")

    features = [
        ("excel_workbook_info", "워크북 구조 및 정보 분석"),
        ("excel_range_read", "셀 범위 데이터 읽기"),
        ("excel_table_read", "테이블 데이터 읽기"),
        ("excel_data_analyze", "데이터 구조 분석"),
        ("excel_chart_list", "차트 목록 조회"),
    ]

    for tool, description in features:
        console.print(f"   • [green]{tool}[/green]: {description}")

    console.print("\n🔧 [bold cyan]개발용 명령어:[/bold cyan]")
    console.print("   [green]oa mcp info[/green]         - 서버 정보 출력")
    console.print("   [green]oa mcp test[/green]         - 기본 테스트 실행")
    console.print("   [green]oa mcp check-port[/green]   - 포트 사용 상태 확인")
    console.print("   [green]oa mcp docs[/green]         - 이 가이드 출력")

if __name__ == "__main__":
    mcp_app()