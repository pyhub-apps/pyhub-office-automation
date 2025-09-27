"""
Email account configuration functionality
Manages email accounts in Windows Credential Manager with app passwords
"""

import json
import sys
from typing import Dict, Optional

import keyring
import typer
from rich.console import Console
from rich.prompt import Prompt
from rich.table import Table

# Provider configurations
PROVIDER_CONFIGS = {
    "gmail": {
        "server": "smtp.gmail.com",
        "port": 587,
        "use_tls": True,
        "guide_url": "https://support.google.com/accounts/answer/185833",
        "name": "Gmail",
    },
    "outlook": {
        "server": "smtp-mail.outlook.com",
        "port": 587,
        "use_tls": True,
        "guide_url": "https://support.microsoft.com/account-billing",
        "name": "Outlook.com",
    },
    "naver": {
        "server": "smtp.naver.com",
        "port": 587,
        "use_tls": True,
        "guide_url": "https://help.naver.com/alias/mail/mail_26.naver",
        "name": "Naver Mail",
    },
    "custom": {"server": None, "port": 587, "use_tls": True, "guide_url": None, "name": "Custom SMTP"},
}


def email_config(
    provider: str = typer.Option(None, "--provider", help="이메일 제공자 (gmail, outlook, naver, custom)"),
    username: str = typer.Option(None, "--username", help="이메일 주소"),
    account_name: str = typer.Option("default", "--account-name", help="계정 별칭 (기본값: default)"),
    server: str = typer.Option(None, "--server", help="SMTP 서버 (custom 제공자용)"),
    port: int = typer.Option(None, "--port", help="SMTP 포트 (기본값: 587)"),
    no_tls: bool = typer.Option(False, "--no-tls", help="TLS 사용 안함"),
) -> Dict:
    """이메일 계정 설정 (앱 비밀번호 기반)"""

    console = Console()

    try:
        import keyring
    except ImportError:
        return {"status": "error", "message": "keyring 라이브러리가 설치되지 않았습니다. 설치: pip install keyring"}

    # Interactive setup if no provider specified
    if not provider:
        provider = _prompt_provider_selection(console)

    if provider not in PROVIDER_CONFIGS:
        return {
            "status": "error",
            "message": f"지원하지 않는 제공자: {provider}. 지원 제공자: {', '.join(PROVIDER_CONFIGS.keys())}",
        }

    config = PROVIDER_CONFIGS[provider]
    console.print(f"\n📧 {config['name']} 계정 설정")

    # Show app password guide
    if config.get("guide_url"):
        console.print(f"\n💡 앱 비밀번호 생성 가이드: {config['guide_url']}")

    # Get username
    if not username:
        username = Prompt.ask("이메일 주소를 입력하세요")

    # Validate email format
    if "@" not in username:
        return {"status": "error", "message": "올바른 이메일 주소 형식이 아닙니다."}

    # Get password securely
    password = typer.prompt("앱 비밀번호를 입력하세요", hide_input=True)

    if not password:
        return {"status": "error", "message": "비밀번호가 입력되지 않았습니다."}

    # Get server settings for custom provider
    if provider == "custom":
        if not server:
            server = Prompt.ask("SMTP 서버 주소를 입력하세요")
        if not port:
            port = typer.prompt("SMTP 포트", default=587, type=int)

    # Use provider defaults
    smtp_server = server or config["server"]
    smtp_port = port or config["port"]
    use_tls = not no_tls and config["use_tls"]

    # Store credentials in Credential Manager
    try:
        service_name = f"oa-email-{account_name}"

        # Store each setting separately
        keyring.set_password(service_name, "username", username)
        keyring.set_password(service_name, "password", password)
        keyring.set_password(service_name, "server", smtp_server)
        keyring.set_password(service_name, "port", str(smtp_port))
        keyring.set_password(service_name, "use_tls", "true" if use_tls else "false")

        console.print(f"\n✅ 계정 '{account_name}'이 성공적으로 설정되었습니다!")

        # Show summary
        _display_account_summary(console, account_name, username, provider, smtp_server, smtp_port, use_tls)

        return {
            "status": "success",
            "message": f"계정 '{account_name}' 설정 완료",
            "account_info": {
                "account_name": account_name,
                "username": username,
                "provider": provider,
                "server": smtp_server,
                "port": smtp_port,
                "use_tls": use_tls,
            },
        }

    except Exception as e:
        return {"status": "error", "message": f"계정 설정 실패: {e}"}


def _prompt_provider_selection(console: Console) -> str:
    """Interactive provider selection"""

    console.print("\n📋 이메일 제공자를 선택하세요:")

    table = Table()
    table.add_column("번호", style="cyan", width=4)
    table.add_column("제공자", style="green")
    table.add_column("설명", style="white")

    providers = list(PROVIDER_CONFIGS.keys())
    for i, provider in enumerate(providers, 1):
        config = PROVIDER_CONFIGS[provider]
        table.add_row(str(i), provider, config["name"])

    console.print(table)

    while True:
        try:
            choice = typer.prompt("\n선택 (번호 입력)", type=int)
            if 1 <= choice <= len(providers):
                selected = providers[choice - 1]
                console.print(f"선택됨: {PROVIDER_CONFIGS[selected]['name']}")
                return selected
            else:
                console.print("❌ 올바른 번호를 입력하세요.")
        except typer.Abort:
            console.print("\n❌ 설정이 취소되었습니다.")
            raise typer.Exit(1)


def _display_account_summary(
    console: Console, account_name: str, username: str, provider: str, server: str, port: int, use_tls: bool
):
    """Display account configuration summary"""

    console.print("\n📊 계정 설정 요약:")

    table = Table()
    table.add_column("항목", style="cyan", width=12)
    table.add_column("값", style="green")

    table.add_row("계정명", account_name)
    table.add_row("이메일", username)
    table.add_row("제공자", provider)
    table.add_row("SMTP 서버", server)
    table.add_row("포트", str(port))
    table.add_row("TLS", "✅" if use_tls else "❌")

    console.print(table)

    console.print("\n💡 사용법:")
    console.print(f'   oa email send --account {account_name} --to recipient@example.com --prompt "메시지 내용"')


if __name__ == "__main__":
    # 직접 실행 시 테스트
    result = email_config()
    print(json.dumps(result, indent=2, ensure_ascii=False))
