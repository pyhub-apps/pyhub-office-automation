"""
Email account management and listing functionality
"""

import json
import sys
from typing import Dict, List, Optional

import keyring
import typer
from rich.console import Console
from rich.table import Table


def list_email_accounts(
    format_output: str = typer.Option("table", "--format", help="출력 형식 (table, json)"),
    verbose: bool = typer.Option(False, "--verbose", help="상세 정보 표시"),
) -> Dict:
    """등록된 이메일 계정 목록 조회"""

    console = Console()

    if verbose:
        console.print("🔍 Windows Credential Manager에서 이메일 계정 검색 중...")

    # oa-email로 시작하는 자격증명 검색
    accounts = _discover_email_accounts(verbose)

    if format_output == "json":
        return {"status": "success", "version": "1.0.0", "accounts": accounts, "total_count": len(accounts)}
    else:
        _display_accounts_table(accounts, console)
        return {"status": "success", "message": f"총 {len(accounts)}개 계정 발견", "accounts": accounts}


def _discover_email_accounts(verbose: bool = False) -> List[Dict]:
    """Windows Credential Manager에서 oa-email 계정들 발견"""

    accounts = []

    try:
        # Windows Credential Manager 접근
        if sys.platform == "win32":
            accounts = _scan_windows_credentials(verbose)
        else:
            # 다른 플랫폼에서는 keyring 기본 백엔드 사용
            accounts = _scan_keyring_credentials(verbose)

    except Exception as e:
        if verbose:
            console = Console()
            console.print(f"❌ 자격증명 검색 오류: {e}")

    return accounts


def _scan_windows_credentials(verbose: bool = False) -> List[Dict]:
    """Windows Credential Manager에서 직접 스캔"""

    accounts = []
    console = Console() if verbose else None

    try:
        import win32con
        import win32cred

        if verbose:
            console.print("   - Windows Credential Manager API 사용")

        # 모든 자격증명 조회
        creds = win32cred.CredEnumerate(None, 0)

        for cred in creds:
            target_name = cred["TargetName"]

            # oa-email로 시작하는 항목 필터링
            if target_name.startswith("oa-email-"):
                account_name = target_name.replace("oa-email-", "").split(":")[0]

                if verbose:
                    console.print(f"   - 발견: {account_name}")

                # 해당 계정의 상세 정보 수집
                account_info = _get_account_details(account_name, verbose)
                if account_info:
                    accounts.append(account_info)

    except ImportError:
        if verbose:
            console.print("   - pywin32 없음, keyring 기본 백엔드 사용")
        accounts = _scan_keyring_credentials(verbose)
    except Exception as e:
        if verbose:
            console.print(f"   - Windows API 오류: {e}, keyring 사용")
        accounts = _scan_keyring_credentials(verbose)

    return accounts


def _scan_keyring_credentials(verbose: bool = False) -> List[Dict]:
    """keyring을 통한 기본 스캔 (제한적)"""

    accounts = []
    console = Console() if verbose else None

    # 일반적인 계정명들을 시도해보기
    common_names = ["default", "gmail", "outlook", "work", "personal", "main"]

    for account_name in common_names:
        try:
            service_name = f"oa-email-{account_name}"
            username = keyring.get_password(service_name, "username")

            if username:  # 계정이 존재함
                if verbose:
                    console.print(f"   - 발견: {account_name}")

                account_info = _get_account_details(account_name, verbose)
                if account_info:
                    accounts.append(account_info)

        except Exception:
            continue  # 해당 계정 없음

    return accounts


def _get_account_details(account_name: str, verbose: bool = False) -> Optional[Dict]:
    """특정 계정의 상세 정보 조회"""

    try:
        service_name = f"oa-email-{account_name}"

        username = keyring.get_password(service_name, "username")
        server = keyring.get_password(service_name, "server")
        port = keyring.get_password(service_name, "port")
        use_tls = keyring.get_password(service_name, "use_tls")

        if not username:
            return None

        # 제공자 추정
        provider = "unknown"
        if server:
            if "gmail" in server:
                provider = "gmail"
            elif "outlook" in server:
                provider = "outlook"
            elif "naver" in server:
                provider = "naver"

        return {
            "account_name": account_name,
            "username": username,
            "provider": provider,
            "server": server or "unknown",
            "port": int(port) if port else 587,
            "use_tls": use_tls == "true" if use_tls else True,
            "status": "configured",
        }

    except Exception as e:
        if verbose:
            console = Console()
            console.print(f"   - {account_name} 상세정보 조회 실패: {e}")
        return None


def _display_accounts_table(accounts: List[Dict], console: Console):
    """계정 목록을 테이블 형태로 출력"""

    if not accounts:
        console.print("📭 등록된 이메일 계정이 없습니다.")
        console.print("\n💡 계정 등록: oa email config --username your@email.com")
        return

    table = Table(title="📧 등록된 이메일 계정")

    table.add_column("계정명", style="cyan", no_wrap=True)
    table.add_column("이메일 주소", style="green")
    table.add_column("제공자", style="blue")
    table.add_column("서버", style="magenta")
    table.add_column("포트", justify="center")
    table.add_column("TLS", justify="center")
    table.add_column("상태", style="yellow")

    for account in accounts:
        table.add_row(
            account["account_name"],
            account["username"],
            account["provider"].upper(),
            account["server"],
            str(account["port"]),
            "✅" if account["use_tls"] else "❌",
            account["status"],
        )

    console.print(table)
    console.print(f"\n📊 총 {len(accounts)}개 계정이 등록되어 있습니다.")


def delete_email_account(
    account_name: str = typer.Argument(..., help="삭제할 계정명"),
    confirm: bool = typer.Option(False, "--confirm", help="확인 없이 삭제"),
) -> Dict:
    """등록된 이메일 계정 삭제"""

    console = Console()

    # 계정 존재 확인
    account_info = _get_account_details(account_name)
    if not account_info:
        return {"status": "error", "message": f"계정 '{account_name}'을 찾을 수 없습니다."}

    # 확인 프로세스
    if not confirm:
        console.print(f"⚠️  계정 '{account_name}' ({account_info['username']})을 삭제하시겠습니까?")
        if not typer.confirm("정말로 삭제하시겠습니까?"):
            return {"status": "cancelled", "message": "삭제가 취소되었습니다."}

    try:
        service_name = f"oa-email-{account_name}"

        # 모든 관련 자격증명 삭제
        for key in ["username", "password", "server", "port", "use_tls"]:
            try:
                keyring.delete_password(service_name, key)
            except keyring.errors.PasswordDeleteError:
                pass  # 이미 없는 경우 무시

        console.print(f"✅ 계정 '{account_name}'이 성공적으로 삭제되었습니다.")

        return {"status": "success", "message": f"계정 '{account_name}' 삭제 완료", "deleted_account": account_info}

    except Exception as e:
        return {"status": "error", "message": f"계정 삭제 실패: {e}"}


if __name__ == "__main__":
    # 직접 실행 시 테스트
    result = list_email_accounts()
    print(json.dumps(result, indent=2, ensure_ascii=False))
