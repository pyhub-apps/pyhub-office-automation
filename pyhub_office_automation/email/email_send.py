"""
AI-powered email sending command
Core email automation functionality with AI content generation
"""

import json
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

import typer
from rich.console import Console
from rich.panel import Panel
from rich.prompt import Confirm, Prompt
from rich.syntax import Syntax

from pyhub_office_automation.version import get_version

from .ai_providers import AIGenerationError, detect_available_providers, generate_email_content
from .email_backends import EmailBackendError, detect_available_backends, send_email, validate_email_backend

console = Console()


def email_send(
    to: str = typer.Option(..., "--to", help="받는 사람 이메일 주소"),
    prompt: Optional[str] = typer.Option(None, "--prompt", help="AI 이메일 생성 프롬프트"),
    prompt_file: Optional[str] = typer.Option(None, "--prompt-file", help="프롬프트를 읽을 파일 경로"),
    subject: Optional[str] = typer.Option(None, "--subject", help="이메일 제목 (AI 생성 시 무시됨)"),
    body: Optional[str] = typer.Option(None, "--body", help="이메일 본문 (AI 생성 시 무시됨)"),
    body_file: Optional[str] = typer.Option(None, "--body-file", help="본문을 읽을 파일 경로"),
    ai_provider: str = typer.Option("auto", "--ai-provider", help="AI 제공자 [auto|claude|codex|gemini|openai|anthropic]"),
    api_key: Optional[str] = typer.Option(None, "--api-key", help="AI API 키 (API 제공자용)"),
    language: str = typer.Option("ko", "--language", help="언어 [ko|en]"),
    tone: str = typer.Option("business", "--tone", help="어조 [formal|casual|business]"),
    from_address: Optional[str] = typer.Option(None, "--from", help="보내는 사람 이메일 주소"),
    cc: Optional[str] = typer.Option(None, "--cc", help="참조 이메일 주소 (쉼표로 구분)"),
    bcc: Optional[str] = typer.Option(None, "--bcc", help="숨은 참조 이메일 주소 (쉼표로 구분)"),
    attachments: Optional[str] = typer.Option(None, "--attachments", help="첨부 파일 경로 (쉼표로 구분)"),
    body_type: str = typer.Option("text", "--body-type", help="본문 형식 [text|html]"),
    priority: str = typer.Option("normal", "--priority", help="이메일 우선순위 [high|normal|low]"),
    backend: str = typer.Option("auto", "--backend", help="이메일 백엔드 [auto|outlook|smtp]"),
    account: Optional[str] = typer.Option(None, "--account", help="사용할 이메일 계정명 (Credential Manager 저장된 계정)"),
    smtp_server: Optional[str] = typer.Option(None, "--smtp-server", help="SMTP 서버 주소"),
    smtp_port: Optional[int] = typer.Option(None, "--smtp-port", help="SMTP 포트"),
    smtp_user: Optional[str] = typer.Option(None, "--smtp-user", help="SMTP 사용자명"),
    smtp_password: Optional[str] = typer.Option(None, "--smtp-password", help="SMTP 비밀번호"),
    confirm: bool = typer.Option(True, "--confirm/--no-confirm", help="발송 전 확인"),
    dry_run: bool = typer.Option(False, "--dry-run", help="실제 발송하지 않고 미리보기만"),
    max_retries: int = typer.Option(3, "--max-retries", help="AI 생성 최대 재시도 횟수"),
    verbose: bool = typer.Option(False, "--verbose", help="상세 내부 동작 과정 출력"),
    output_format: str = typer.Option("json", "--format", help="출력 형식 [json|text]"),
):
    """
    AI 기반 이메일 생성 및 발송

    AI 프롬프트를 사용해 이메일 내용을 자동 생성하거나 수동으로 제목/본문 지정 가능

    예제:
        oa email send --to "user@example.com" --prompt "회의 일정 변경 안내"
        oa email send --to "user@example.com" --prompt-file "email_prompt.txt"
        oa email send --to "client@company.com" --subject "안녕하세요" --body "테스트 메일입니다"
    """

    start_time = time.time()

    try:
        if verbose:
            console.print("🔍 [bold blue]이메일 발송 프로세스 시작[/bold blue]")
            console.print(f"   • 받는 사람: {to}")
            console.print(f"   • AI 프롬프트: {prompt if prompt else '없음'}")
            console.print(f"   • 프롬프트 파일: {prompt_file if prompt_file else '없음'}")
            console.print(f"   • 계정: {account if account else '기본값'}")
            console.print(f"   • 백엔드: {backend}")
            console.print(f"   • 언어: {language}, 어조: {tone}")

        # Validate inputs
        if verbose:
            console.print("\n📋 [bold yellow]Step 1: 입력 파라미터 검증[/bold yellow]")
        _validate_inputs(to, prompt, prompt_file, subject, body, body_file, verbose)

        # Step 2: Early email backend validation (before AI generation)
        if verbose:
            console.print("\n📧 [bold yellow]Step 2: 이메일 발송 가능성 사전 검증[/bold yellow]")

        # Prepare SMTP config early for validation
        smtp_config = None
        if backend in ["smtp", "auto"]:
            if verbose:
                console.print("   • SMTP 설정 준비 중...")
            smtp_config = _prepare_smtp_config(smtp_server, smtp_port, smtp_user, smtp_password, account, verbose)

        # Validate email backend before AI generation
        backend_validation = validate_email_backend(backend, smtp_config, verbose)

        if backend_validation["status"] == "error":
            return _create_response(
                "error",
                {"error": f"이메일 백엔드 검증 실패: {backend_validation['message']}", "backend_error": backend_validation},
            )
        elif backend_validation["status"] == "warning":
            if verbose:
                console.print(f"   ⚠️  경고: {backend_validation['message']}")
            # Continue with warning but inform user

        if verbose:
            console.print(f"   • ✅ 이메일 백엔드 '{backend_validation['backend']}' 검증 완료")

        # Step 3: Generate or prepare email content
        if verbose:
            console.print("\n✍️ [bold yellow]Step 3: 이메일 내용 준비[/bold yellow]")
        email_content = _prepare_email_content(
            prompt, prompt_file, subject, body, body_file, ai_provider, api_key, language, tone, to, max_retries, verbose
        )

        final_subject = email_content["subject"]
        final_body = email_content["body"]
        ai_generation_info = email_content.get("ai_info")

        # Step 4: Prepare attachments
        if verbose:
            console.print("\n📎 [bold yellow]Step 4: 첨부 파일 처리[/bold yellow]")
        attachment_list = _prepare_attachments(attachments, verbose)

        # Step 5: Show confirmation if needed
        if verbose:
            console.print("\n✅ [bold yellow]Step 5: 발송 확인[/bold yellow]")
        if confirm and not dry_run:
            confirmed = _show_confirmation(
                to, final_subject, final_body, from_address, cc, bcc, attachment_list, body_type, verbose
            )
            if not confirmed:
                return _create_response(
                    "cancelled", {"message": "사용자가 발송을 취소했습니다", "ai_generation": ai_generation_info}
                )

        # Step 6: Send email (or dry run)
        if verbose:
            console.print("\n📧 [bold yellow]Step 6: 이메일 발송[/bold yellow]")
        if dry_run:
            if verbose:
                console.print("   • 드라이런 모드 - 실제 발송하지 않음")
            _show_dry_run_preview(to, final_subject, final_body, from_address, cc, bcc, attachment_list, body_type, verbose)
            return _create_response(
                "preview",
                {
                    "message": "미리보기 모드 - 실제 발송되지 않음",
                    "email_preview": {"to": to, "subject": final_subject, "body": final_body, "attachments": attachment_list},
                    "ai_generation": ai_generation_info,
                },
            )

        # Use validated backend and SMTP config from earlier step
        final_backend = backend_validation["backend"]

        # Send email
        if verbose:
            console.print(f"   • 백엔드 '{final_backend}'로 이메일 발송 중...")
        send_result = send_email(
            to=to,
            subject=final_subject,
            body=final_body,
            from_address=from_address,
            cc=cc,
            bcc=bcc,
            attachments=attachment_list,
            body_type=body_type,
            backend=final_backend,
            smtp_config=smtp_config,
            account_name=account,
        )

        if verbose:
            console.print(f"   • ✅ 이메일 발송 완료 (백엔드: {send_result['backend']})")
            if send_result.get("message_id"):
                console.print(f"   • 메시지 ID: {send_result['message_id']}")

        execution_time = round((time.time() - start_time) * 1000)

        return _create_response(
            "success",
            {
                "email": {
                    "to": send_result["to"],
                    "cc": send_result.get("cc", []),
                    "bcc": send_result.get("bcc", []),
                    "subject": final_subject,
                    "body_type": body_type,
                    "attachments": attachment_list,
                    "sent_at": datetime.now().isoformat(),
                    "backend": send_result["backend"],
                    "message_id": send_result.get("message_id"),
                },
                "ai_generation": ai_generation_info,
                "execution_time_ms": execution_time,
            },
        )

    except (AIGenerationError, EmailBackendError) as e:
        return _create_response("error", {"error": str(e)})
    except Exception as e:
        return _create_response("error", {"error": f"Unexpected error: {e}"})


def _validate_inputs(
    to: str,
    prompt: Optional[str],
    prompt_file: Optional[str],
    subject: Optional[str],
    body: Optional[str],
    body_file: Optional[str],
    verbose: bool = False,
):
    """Validate input parameters"""

    if verbose:
        console.print("   • 이메일 주소 형식 검증...")
    # Email validation (basic)
    if "@" not in to:
        raise ValueError("유효하지 않은 이메일 주소입니다")

    if verbose:
        console.print("   • 콘텐츠 소스 검증...")
    # Content source validation
    content_sources = sum([bool(prompt), bool(prompt_file), bool(subject and body), bool(body_file)])
    if content_sources == 0:
        raise ValueError("--prompt, --prompt-file, --subject + --body, 또는 --body-file 중 하나를 지정해야 합니다")

    if content_sources > 1:
        raise ValueError("--prompt, --prompt-file, --subject + --body, --body-file 중 하나만 지정할 수 있습니다")

    if verbose:
        console.print("   • 파일 경로 검증...")
    # File validation
    if body_file and not Path(body_file).exists():
        raise FileNotFoundError(f"본문 파일을 찾을 수 없습니다: {body_file}")

    if prompt_file and not Path(prompt_file).exists():
        raise FileNotFoundError(f"프롬프트 파일을 찾을 수 없습니다: {prompt_file}")

    if verbose:
        console.print("   • ✅ 모든 입력 검증 완료")


def _prepare_email_content(
    prompt: Optional[str],
    prompt_file: Optional[str],
    subject: Optional[str],
    body: Optional[str],
    body_file: Optional[str],
    ai_provider: str,
    api_key: Optional[str],
    language: str,
    tone: str,
    recipient: str,
    max_retries: int,
    verbose: bool = False,
) -> Dict:
    """Prepare email content (AI generation or manual)"""

    # Handle prompt from file
    final_prompt = prompt
    if prompt_file:
        if verbose:
            console.print(f"   • 프롬프트 파일에서 읽기: {prompt_file}")
        with open(prompt_file, "r", encoding="utf-8") as f:
            final_prompt = f.read().strip()
        if verbose:
            console.print(f"   • 파일에서 읽은 프롬프트: '{final_prompt[:50]}{'...' if len(final_prompt) > 50 else ''}'")

    if final_prompt:
        # AI generation
        if verbose:
            console.print(f"   • AI 제공자 검증 및 선택: {ai_provider}")
            console.print(f"   • 프롬프트 내용: '{final_prompt[:50]}{'...' if len(final_prompt) > 50 else ''}'")

        console.print(f"🤖 AI로 이메일 내용 생성 중... (Provider: {ai_provider})")

        ai_content = generate_email_content(
            prompt=final_prompt,
            provider=ai_provider,
            api_key=api_key,
            language=language,
            tone=tone,
            recipient=recipient,
            max_retries=max_retries,
            verbose=verbose,
        )

        if verbose:
            console.print(
                f"   • ✅ AI 생성 완료 - 제목: '{ai_content['subject'][:30]}{'...' if len(ai_content['subject']) > 30 else ''}'"
            )
            console.print(f"   • 본문 길이: {len(ai_content['body'])}자")

        return {
            "subject": ai_content["subject"],
            "body": ai_content["body"],
            "ai_info": {
                "provider": ai_provider,
                "prompt": final_prompt,
                "prompt_source": "file" if prompt_file else "direct",
                "language": language,
                "tone": tone,
                "generated_at": datetime.now().isoformat(),
            },
        }

    elif body_file:
        # Read from file
        if verbose:
            console.print(f"   • 파일에서 본문 읽기: {body_file}")
        with open(body_file, "r", encoding="utf-8") as f:
            file_content = f.read()

        # Use first line as subject if not provided
        lines = file_content.split("\n")
        file_subject = subject or (lines[0] if lines else "No Subject")
        file_body = "\n".join(lines[1:]) if len(lines) > 1 else file_content

        if verbose:
            console.print(f"   • ✅ 파일 읽기 완료 - {len(lines)}줄, {len(file_content)}자")

        return {"subject": file_subject, "body": file_body, "ai_info": None}

    else:
        # Manual subject/body
        if verbose:
            console.print("   • 수동 제목/본문 사용")
            console.print(f"   • 제목: '{subject}'")
            console.print(f"   • 본문 길이: {len(body)}자")

        return {"subject": subject, "body": body, "ai_info": None}


def _prepare_attachments(attachments: Optional[str], verbose: bool = False) -> List[str]:
    """Prepare attachment file list"""
    if not attachments:
        if verbose:
            console.print("   • 첨부 파일 없음")
        return []

    if verbose:
        console.print("   • 첨부 파일 경로 처리 중...")

    attachment_list = []
    for path in attachments.split(","):
        path = path.strip()
        if path:
            if verbose:
                console.print(f"     - 파일 확인: {path}")
            path_obj = Path(path)
            if not path_obj.exists():
                raise FileNotFoundError(f"첨부 파일을 찾을 수 없습니다: {path}")
            attachment_list.append(str(path_obj.resolve()))
            if verbose:
                file_size = path_obj.stat().st_size
                console.print(f"       ✅ 크기: {file_size:,} bytes")

    if verbose:
        console.print(f"   • ✅ 첨부 파일 {len(attachment_list)}개 준비 완료")

    return attachment_list


def _show_confirmation(
    to: str,
    subject: str,
    body: str,
    from_address: Optional[str],
    cc: Optional[str],
    bcc: Optional[str],
    attachments: List[str],
    body_type: str,
    verbose: bool = False,
) -> bool:
    """Show interactive confirmation dialog"""

    if verbose:
        console.print("   • 사용자 확인 대화창 표시")

    console.print("\n" + "=" * 60)
    console.print("📧 이메일 발송 확인", style="bold blue")
    console.print("=" * 60)

    # Email details
    console.print(f"받는 사람: {to}", style="cyan")
    if from_address:
        console.print(f"보내는 사람: {from_address}", style="dim")
    if cc:
        console.print(f"참조: {cc}", style="dim")
    if bcc:
        console.print(f"숨은 참조: {bcc}", style="dim")

    console.print(f"\n제목: {subject}", style="yellow")

    # Show body preview
    console.print(f"\n본문 ({body_type}):", style="green")
    if body_type == "html":
        syntax = Syntax(body, "html", theme="github-dark", line_numbers=False)
        console.print(Panel(syntax, title="HTML 본문"))
    else:
        console.print(Panel(body, title="텍스트 본문"))

    # Show attachments
    if attachments:
        console.print(f"\n첨부 파일:", style="magenta")
        for attachment in attachments:
            console.print(f"  • {Path(attachment).name}")

    console.print("=" * 60)

    # Confirmation prompt
    choices = ["y", "n", "e"]
    response = Prompt.ask("\n이 이메일을 발송하시겠습니까?", choices=choices, default="y", show_choices=True, console=console)

    if response == "e":
        # Edit functionality
        console.print("✏️  편집 기능은 향후 구현 예정입니다", style="yellow")
        return Confirm.ask("편집 없이 발송하시겠습니까?")

    return response == "y"


def _show_dry_run_preview(
    to: str,
    subject: str,
    body: str,
    from_address: Optional[str],
    cc: Optional[str],
    bcc: Optional[str],
    attachments: List[str],
    body_type: str,
    verbose: bool = False,
):
    """Show dry run preview"""

    if verbose:
        console.print("   • 드라이런 미리보기 표시")

    console.print("\n" + "=" * 60)
    console.print("🔍 미리보기 모드 (실제 발송 안 됨)", style="bold yellow")
    console.print("=" * 60)

    console.print(f"받는 사람: {to}")
    if from_address:
        console.print(f"보내는 사람: {from_address}")
    if cc:
        console.print(f"참조: {cc}")
    if bcc:
        console.print(f"숨은 참조: {bcc}")

    console.print(f"\n제목: {subject}")
    console.print(f"\n본문:\n{body}")

    if attachments:
        console.print(f"\n첨부 파일:")
        for attachment in attachments:
            console.print(f"  • {attachment}")


def _prepare_smtp_config(
    server: Optional[str],
    port: Optional[int],
    user: Optional[str],
    password: Optional[str],
    account: Optional[str] = None,
    verbose: bool = False,
) -> Optional[Dict]:
    """Prepare SMTP configuration"""
    import os

    from .email_backends import get_account_smtp_config

    # 계정이 지정된 경우 해당 계정 설정 우선 사용
    if account:
        if verbose:
            console.print(f"     - 계정 '{account}' 설정 사용")

        account_config = get_account_smtp_config(account)
        if account_config:
            if verbose:
                console.print(f"     - 계정 설정 로드 성공: {account_config['server']}:{account_config['port']}")
            return account_config
        else:
            if verbose:
                console.print(f"     - 계정 '{account}' 설정을 찾을 수 없음, 기본 설정 사용")

    # 기본 SMTP 설정 (환경변수 또는 매개변수)
    config = {
        "server": server or os.getenv("SMTP_SERVER", "smtp.gmail.com"),
        "port": port or int(os.getenv("SMTP_PORT", "587")),
        "username": user or os.getenv("SMTP_USERNAME"),
        "password": password or os.getenv("SMTP_PASSWORD"),
        "use_tls": os.getenv("SMTP_USE_TLS", "true").lower() == "true",
    }

    if verbose:
        console.print(f"     - SMTP 서버: {config['server']}:{config['port']}")
        console.print(f"     - TLS 사용: {config['use_tls']}")
        console.print(f"     - 사용자명: {config['username'] or '환경변수에서 읽기'}")
        console.print(f"     - 비밀번호: {'설정됨' if config['password'] else '환경변수에서 읽기'}")

    return config


def _create_response(status: str, data: Dict) -> None:
    """Create and output structured response"""
    response = {
        "status": status,
        "command": "email-send",
        "version": get_version(),
        "data": data,
        "timestamp": datetime.now().isoformat(),
    }

    if status == "success":
        response["ai_summary"] = f"이메일이 성공적으로 발송되었습니다: {data['email']['to']}"
    elif status == "cancelled":
        response["ai_summary"] = "이메일 발송이 취소되었습니다"
    elif status == "preview":
        response["ai_summary"] = "이메일 미리보기가 완료되었습니다 (실제 발송 안 됨)"
    elif status == "error":
        response["ai_summary"] = f"이메일 발송 중 오류 발생: {data.get('error', 'Unknown error')}"

    # Output JSON
    try:
        json_output = json.dumps(response, ensure_ascii=False, indent=2)
        typer.echo(json_output)
    except UnicodeEncodeError:
        json_output = json.dumps(response, ensure_ascii=True, indent=2)
        typer.echo(json_output)
