"""
Email backend implementations for different platforms
Supports Outlook COM (Windows) and SMTP (cross-platform)
"""

import os
import smtplib
import sys
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Dict, List, Optional, Union

try:
    import keyring

    KEYRING_AVAILABLE = True
except ImportError:
    KEYRING_AVAILABLE = False


class EmailBackendError(Exception):
    """Email backend error"""

    pass


def detect_available_backends() -> Dict[str, bool]:
    """Detect which email backends are available"""
    backends = {"outlook": False, "smtp": True}  # SMTP is always available

    # Check Outlook COM availability (Windows only)
    if sys.platform == "win32":
        try:
            import win32com.client

            backends["outlook"] = True
        except ImportError:
            pass

    return backends


def validate_email_backend(backend: str, smtp_config: Optional[Dict] = None, verbose: bool = False) -> Dict[str, str]:
    """Validate email backend configuration and availability"""

    if verbose:
        from rich.console import Console

        console = Console()
        console.print(f"       - 이메일 백엔드 '{backend}' 검증 중...")

    # Auto-detect backend
    if backend == "auto":
        available = detect_available_backends()

        # 기본 계정이 있는지 확인하여 SMTP 우선 사용
        default_config = get_account_smtp_config("default")
        if default_config:
            backend = "smtp"
        else:
            backend = "outlook" if available["outlook"] else "smtp"

        if verbose:
            console.print(f"       - 자동 선택된 백엔드: {backend}")

    # Backend-specific validation
    if backend == "outlook":
        return _validate_outlook_backend(verbose)
    elif backend == "smtp":
        return _validate_smtp_backend(smtp_config, verbose)
    else:
        return {"status": "error", "message": f"지원하지 않는 백엔드: {backend}", "backend": backend}


def _validate_outlook_backend(verbose: bool = False) -> Dict[str, str]:
    """Validate Outlook COM backend"""

    if verbose:
        from rich.console import Console

        console = Console()

    # Check Windows platform
    if sys.platform != "win32":
        return {"status": "error", "message": "Outlook은 Windows에서만 사용 가능합니다", "backend": "outlook"}

    # Check COM library
    try:
        import win32com.client

        if verbose:
            console.print("         - pywin32 라이브러리 확인됨")
    except ImportError:
        return {
            "status": "error",
            "message": "pywin32 라이브러리가 설치되지 않았습니다. 설치: pip install pywin32",
            "backend": "outlook",
        }

    # Quick availability check (don't test actual connection)
    if verbose:
        console.print("         - Outlook 기본 요구사항 확인됨")

    return {"status": "success", "message": "Outlook COM 백엔드 사용 가능 (실제 연결은 발송 시 확인)", "backend": "outlook"}


def _validate_smtp_backend(smtp_config: Optional[Dict], verbose: bool = False) -> Dict[str, str]:
    """Validate SMTP backend configuration"""

    if verbose:
        from rich.console import Console

        console = Console()

    if not smtp_config:
        # Load default SMTP config
        smtp_config = _get_default_smtp_config()

    # Check required SMTP settings
    required_fields = ["server", "port", "username", "password"]
    missing_fields = []

    for field in required_fields:
        if not smtp_config.get(field):
            missing_fields.append(field)

    if missing_fields:
        if verbose:
            console.print(f"         - SMTP 설정 누락: {', '.join(missing_fields)}")
        return {
            "status": "warning",
            "message": f"SMTP 설정이 누락되었습니다: {', '.join(missing_fields)}. 환경변수 또는 명령어 옵션으로 설정하세요.",
            "backend": "smtp",
            "missing_fields": missing_fields,
        }

    # Test SMTP connection (optional - can be slow)
    if verbose:
        console.print(f"         - SMTP 서버: {smtp_config['server']}:{smtp_config['port']}")
        console.print(f"         - 사용자: {smtp_config['username']}")

    return {
        "status": "success",
        "message": f"SMTP 백엔드 설정 완료 ({smtp_config['server']}:{smtp_config['port']})",
        "backend": "smtp",
    }


def send_email(
    to: Union[str, List[str]],
    subject: str,
    body: str,
    from_address: Optional[str] = None,
    cc: Optional[Union[str, List[str]]] = None,
    bcc: Optional[Union[str, List[str]]] = None,
    attachments: Optional[List[str]] = None,
    body_type: str = "text",
    backend: str = "auto",
    smtp_config: Optional[Dict] = None,
    account_name: Optional[str] = None,
) -> Dict:
    """Send email using specified backend"""

    # Auto-detect backend
    if backend == "auto":
        available = detect_available_backends()

        # 계정이 지정되어 있으면 SMTP 우선 사용
        if account_name:
            account_config = get_account_smtp_config(account_name)
            if account_config:
                backend = "smtp"
            else:
                backend = "outlook" if available["outlook"] else "smtp"
        else:
            # 기본 계정이 있는지 확인
            default_config = get_account_smtp_config("default")
            if default_config:
                backend = "smtp"
            else:
                backend = "outlook" if available["outlook"] else "smtp"

    # Validate backend availability
    available = detect_available_backends()
    if backend not in available or not available[backend]:
        raise EmailBackendError(f"Backend '{backend}' not available")

    # Normalize recipients
    to_list = _normalize_recipients(to)
    cc_list = _normalize_recipients(cc) if cc else []
    bcc_list = _normalize_recipients(bcc) if bcc else []

    # Validate attachments
    if attachments:
        _validate_attachments(attachments)

    # Send using specified backend
    if backend == "outlook":
        return _send_via_outlook(to_list, subject, body, from_address, cc_list, bcc_list, attachments, body_type)
    elif backend == "smtp":
        if not smtp_config:
            # 계정명이 지정된 경우 해당 계정 설정 사용
            if account_name:
                smtp_config = get_account_smtp_config(account_name)
                if not smtp_config:
                    raise EmailBackendError(f"Account '{account_name}' not found or not configured")
            else:
                smtp_config = _get_default_smtp_config()
        return _send_via_smtp(to_list, subject, body, from_address, cc_list, bcc_list, attachments, body_type, smtp_config)
    else:
        raise EmailBackendError(f"Unsupported backend: {backend}")


def _send_via_outlook(
    to_list: List[str],
    subject: str,
    body: str,
    from_address: Optional[str],
    cc_list: List[str],
    bcc_list: List[str],
    attachments: Optional[List[str]],
    body_type: str,
) -> Dict:
    """Send email via Outlook COM"""

    try:
        import win32com.client

        # Create Outlook application
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # 0 = Mail item

        # Set recipients
        mail.To = "; ".join(to_list)
        if cc_list:
            mail.CC = "; ".join(cc_list)
        if bcc_list:
            mail.BCC = "; ".join(bcc_list)

        # Set subject and body
        mail.Subject = subject
        if body_type == "html":
            mail.HTMLBody = body
        else:
            mail.Body = body

        # Add attachments
        if attachments:
            for attachment_path in attachments:
                mail.Attachments.Add(str(Path(attachment_path).resolve()))

        # Send email
        mail.Send()

        return {
            "status": "sent",
            "backend": "outlook",
            "message_id": None,  # Outlook doesn't return message ID
            "to": to_list,
            "cc": cc_list,
            "bcc": bcc_list,
        }

    except ImportError:
        raise EmailBackendError("Outlook COM not available (install pywin32)")
    except Exception as e:
        raise EmailBackendError(f"Outlook error: {e}")


def _send_via_smtp(
    to_list: List[str],
    subject: str,
    body: str,
    from_address: Optional[str],
    cc_list: List[str],
    bcc_list: List[str],
    attachments: Optional[List[str]],
    body_type: str,
    smtp_config: Dict,
) -> Dict:
    """Send email via SMTP"""

    try:
        # Create message
        msg = MIMEMultipart()
        msg["From"] = from_address or smtp_config.get("username")
        msg["To"] = ", ".join(to_list)
        if cc_list:
            msg["Cc"] = ", ".join(cc_list)
        msg["Subject"] = subject

        # Add body
        if body_type == "html":
            msg.attach(MIMEText(body, "html", "utf-8"))
        else:
            msg.attach(MIMEText(body, "plain", "utf-8"))

        # Add attachments
        if attachments:
            for attachment_path in attachments:
                _add_attachment(msg, attachment_path)

        # Connect to SMTP server
        server = smtplib.SMTP(smtp_config["server"], smtp_config["port"])

        if smtp_config.get("use_tls", True):
            server.starttls()

        if smtp_config.get("username") and smtp_config.get("password"):
            server.login(smtp_config["username"], smtp_config["password"])

        # Send email
        all_recipients = to_list + cc_list + bcc_list
        text = msg.as_string()
        server.sendmail(msg["From"], all_recipients, text)
        server.quit()

        return {
            "status": "sent",
            "backend": "smtp",
            "message_id": msg.get("Message-ID"),
            "to": to_list,
            "cc": cc_list,
            "bcc": bcc_list,
            "smtp_server": smtp_config["server"],
        }

    except Exception as e:
        raise EmailBackendError(f"SMTP error: {e}")


def _normalize_recipients(recipients: Union[str, List[str]]) -> List[str]:
    """Normalize recipients to list format"""
    if isinstance(recipients, str):
        # Split by comma or semicolon
        return [email.strip() for email in recipients.replace(";", ",").split(",") if email.strip()]
    elif isinstance(recipients, list):
        return [email.strip() for email in recipients if email.strip()]
    else:
        return []


def _validate_attachments(attachments: List[str]):
    """Validate attachment files exist"""
    for attachment in attachments:
        path = Path(attachment)
        if not path.exists():
            raise EmailBackendError(f"Attachment not found: {attachment}")
        if not path.is_file():
            raise EmailBackendError(f"Attachment is not a file: {attachment}")


def _add_attachment(msg: MIMEMultipart, attachment_path: str):
    """Add attachment to email message"""
    try:
        path = Path(attachment_path)

        with open(path, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())

        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename= {path.name}")
        msg.attach(part)

    except Exception as e:
        raise EmailBackendError(f"Failed to attach {attachment_path}: {e}")


def get_account_smtp_config(account_name: str = "default") -> Optional[Dict]:
    """Get SMTP configuration from Credential Manager by account name"""

    if not KEYRING_AVAILABLE:
        return None

    try:
        service_name = f"oa-email-{account_name}"

        username = keyring.get_password(service_name, "username")
        password = keyring.get_password(service_name, "password")
        server = keyring.get_password(service_name, "server")
        port = keyring.get_password(service_name, "port")
        use_tls = keyring.get_password(service_name, "use_tls")

        if not username or not password:
            return None

        return {
            "server": server or "smtp.gmail.com",
            "port": int(port) if port else 587,
            "username": username,
            "password": password,
            "use_tls": use_tls == "true" if use_tls else True,
        }

    except Exception:
        return None


def list_available_accounts() -> List[str]:
    """List all available email accounts in Credential Manager"""

    if not KEYRING_AVAILABLE:
        return []

    accounts = []

    try:
        # Windows에서 직접 검색
        if sys.platform == "win32":
            try:
                import win32cred

                creds = win32cred.CredEnumerate(None, 0)

                for cred in creds:
                    target_name = cred["TargetName"]
                    if target_name.startswith("oa-email-"):
                        account_name = target_name.replace("oa-email-", "").split(":")[0]
                        if account_name not in accounts:
                            accounts.append(account_name)

            except ImportError:
                # pywin32 없으면 기본 계정들만 체크
                pass

        # 일반적인 계정명들 체크
        common_names = ["default", "gmail", "outlook", "work", "personal", "main"]
        for name in common_names:
            if name not in accounts:
                config = get_account_smtp_config(name)
                if config:
                    accounts.append(name)

    except Exception:
        pass

    return accounts


def _get_default_smtp_config() -> Dict:
    """Get default SMTP configuration from environment variables or Credential Manager"""

    # 먼저 Credential Manager에서 기본 계정 확인
    if KEYRING_AVAILABLE:
        config = get_account_smtp_config("default")
        if config:
            return config

    # 환경변수에서 폴백
    return {
        "server": os.getenv("SMTP_SERVER", "smtp.gmail.com"),
        "port": int(os.getenv("SMTP_PORT", "587")),
        "username": os.getenv("SMTP_USERNAME"),
        "password": os.getenv("SMTP_PASSWORD"),
        "use_tls": os.getenv("SMTP_USE_TLS", "true").lower() == "true",
    }
