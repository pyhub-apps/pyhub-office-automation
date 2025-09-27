"""
AI provider integration for email content generation
Supports both API-based and CLI-based AI providers
"""

import json
import os
import subprocess
import sys
from typing import Dict, Optional, Tuple


class AIGenerationError(Exception):
    """AI content generation error"""

    pass


def detect_available_providers() -> Dict[str, bool]:
    """Detect which AI providers are available on the system"""
    providers = {"claude": False, "codex": False, "gemini": False, "openai": False, "anthropic": False}

    # Check CLI tools
    for cli_tool in ["claude", "codex", "gemini"]:
        try:
            # Windows: PowerShell을 통한 실행 시도
            if sys.platform == "win32":
                result = subprocess.run(
                    ["powershell.exe", "-ExecutionPolicy", "Bypass", "-Command", f"{cli_tool} --version"],
                    capture_output=True,
                    text=True,
                    timeout=5,
                    encoding="utf-8",
                    errors="replace",
                )
            else:
                result = subprocess.run([cli_tool, "--version"], capture_output=True, text=True, timeout=5)

            if result.returncode == 0:
                providers[cli_tool] = True
        except (subprocess.TimeoutExpired, FileNotFoundError):
            pass

    # Check API providers (check if API keys exist)
    if os.getenv("OPENAI_API_KEY"):
        providers["openai"] = True
    if os.getenv("ANTHROPIC_API_KEY"):
        providers["anthropic"] = True

    return providers


def generate_email_content(
    prompt: str,
    provider: str = "auto",
    api_key: Optional[str] = None,
    language: str = "ko",
    tone: str = "business",
    recipient: str = "",
    max_retries: int = 3,
    verbose: bool = False,
) -> Dict[str, str]:
    """Generate email subject and body using AI"""

    if verbose:
        from rich.console import Console

        console = Console()

    # Auto-detect provider if needed
    if provider == "auto":
        if verbose:
            console.print("     - AI 제공자 자동 감지 중...")
        available = detect_available_providers()
        if verbose:
            console.print(f"     - 사용 가능한 제공자: {[k for k, v in available.items() if v]}")
        # Prefer CLI tools first, then API
        for preferred in ["claude", "gemini", "codex", "openai", "anthropic"]:
            if available[preferred]:
                provider = preferred
                if verbose:
                    console.print(f"     - 선택된 제공자: {provider}")
                break
        else:
            raise AIGenerationError("No AI provider available")

    # Validate provider availability
    available = detect_available_providers()
    if provider not in available or not available[provider]:
        if provider in ["openai", "anthropic"] and not api_key:
            raise AIGenerationError(f"{provider} requires API key")
        elif provider in ["claude", "codex", "gemini"]:
            raise AIGenerationError(f"{provider} CLI tool not found")

    if verbose:
        console.print(f"     - 제공자 '{provider}' 검증 완료")

    # Generate content with retries
    last_error = None
    for attempt in range(max_retries):
        try:
            if verbose:
                console.print(f"     - 시도 {attempt + 1}/{max_retries}: AI 콘텐츠 생성 중...")

            if provider in ["openai", "anthropic"]:
                result = _generate_with_api(provider, prompt, api_key, language, tone, recipient, verbose)
            elif provider in ["claude", "codex", "gemini"]:
                result = _generate_with_cli(provider, prompt, language, tone, recipient, verbose)
            else:
                raise AIGenerationError(f"Unsupported provider: {provider}")

            if verbose:
                console.print(f"     - ✅ AI 생성 성공 (시도 {attempt + 1})")
            return result

        except Exception as e:
            last_error = e
            if verbose:
                console.print(f"     - ❌ 시도 {attempt + 1} 실패: {str(e)[:50]}...")
            if attempt < max_retries - 1:
                continue
            else:
                raise AIGenerationError(f"Failed after {max_retries} attempts: {last_error}")


def _generate_with_api(
    provider: str, prompt: str, api_key: str, language: str, tone: str, recipient: str, verbose: bool = False
) -> Dict[str, str]:
    """Generate content using API providers"""

    if verbose:
        from rich.console import Console

        console = Console()
        console.print(f"       - API 프롬프트 구성 중 ({language}, {tone})")

    system_prompt = _build_system_prompt(language, tone, recipient)

    if verbose:
        console.print(f"       - {provider.upper()} API 호출 중...")

    if provider == "openai":
        return _call_openai_api(system_prompt, prompt, api_key, verbose)
    elif provider == "anthropic":
        return _call_anthropic_api(system_prompt, prompt, api_key, verbose)
    else:
        raise AIGenerationError(f"Unsupported API provider: {provider}")


def _generate_with_cli(
    provider: str, prompt: str, language: str, tone: str, recipient: str, verbose: bool = False
) -> Dict[str, str]:
    """Generate content using CLI providers"""

    if verbose:
        from rich.console import Console

        console = Console()
        console.print(f"       - CLI 프롬프트 구성 중 ({language}, {tone})")

    system_prompt = _build_system_prompt(language, tone, recipient)
    combined_prompt = f"{system_prompt}\n\nTask: {prompt}"

    try:
        if provider == "claude":
            cmd = ["claude", "-p", combined_prompt]
        elif provider == "codex":
            cmd = ["codex", "exec", combined_prompt]
        elif provider == "gemini":
            cmd = ["gemini", "-p", combined_prompt]
        else:
            raise AIGenerationError(f"Unsupported CLI provider: {provider}")

        if verbose:
            console.print(f"       - {provider} CLI 실행 중: {' '.join(cmd[:2])}...")

        # Windows에서 PowerShell을 통해 CLI 도구 실행
        if sys.platform == "win32":
            # PowerShell에서 문자열 인자를 안전하게 전달하기 위해 따옴표 처리
            safe_prompt = combined_prompt.replace('"', '""').replace("'", "''")
            ps_cmd = ["powershell.exe", "-ExecutionPolicy", "Bypass", "-Command", f"{provider} -p '{safe_prompt}'"]
            result = subprocess.run(
                ps_cmd,
                capture_output=True,
                text=True,
                timeout=30,
                encoding="utf-8",
                errors="replace",  # 인코딩 오류 시 대체 문자 사용
            )
        else:
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30, encoding="utf-8")

        if result.returncode != 0:
            raise AIGenerationError(f"{provider} CLI error: {result.stderr}")

        if verbose:
            console.print(f"       - CLI 응답 파싱 중 ({len(result.stdout)} chars)")

        return _parse_ai_response(result.stdout, provider, verbose)

    except subprocess.TimeoutExpired:
        raise AIGenerationError(f"{provider} CLI timeout")
    except FileNotFoundError:
        raise AIGenerationError(f"{provider} CLI not found")


def _build_system_prompt(language: str, tone: str, recipient: str) -> str:
    """Build system prompt for AI generation"""

    lang_instruction = {"ko": "한국어로 이메일을 작성하세요.", "en": "Write the email in English."}.get(
        language, "한국어로 이메일을 작성하세요."
    )

    tone_instruction = {
        "formal": "격식 있고 공손한 어조로",
        "casual": "친근하고 편안한 어조로",
        "business": "비즈니스 전문적인 어조로",
    }.get(tone, "비즈니스 전문적인 어조로")

    return f"""
{lang_instruction} {tone_instruction} 작성해주세요.
{f"받는 사람: {recipient}" if recipient else ""}

다음 JSON 형식으로만 응답하세요:
{{
    "subject": "이메일 제목",
    "body": "이메일 본문"
}}

중요: JSON 형식 외에 다른 텍스트는 포함하지 마세요.
"""


def _call_openai_api(system_prompt: str, user_prompt: str, api_key: str, verbose: bool = False) -> Dict[str, str]:
    """Call OpenAI API for content generation"""
    try:
        import openai

        openai.api_key = api_key

        if verbose:
            from rich.console import Console

            console = Console()
            console.print("         - OpenAI API 요청 전송 중...")

        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "system", "content": system_prompt}, {"role": "user", "content": user_prompt}],
            max_tokens=500,
            temperature=0.7,
        )

        content = response.choices[0].message.content

        if verbose:
            console.print(f"         - API 응답 수신: {len(content)} chars")

        return _parse_ai_response(content, "openai", verbose)

    except ImportError:
        raise AIGenerationError("OpenAI library not installed. Run: pip install openai")
    except Exception as e:
        raise AIGenerationError(f"OpenAI API error: {e}")


def _call_anthropic_api(system_prompt: str, user_prompt: str, api_key: str, verbose: bool = False) -> Dict[str, str]:
    """Call Anthropic API for content generation"""
    try:
        import anthropic

        if verbose:
            from rich.console import Console

            console = Console()
            console.print("         - Anthropic API 요청 전송 중...")

        client = anthropic.Anthropic(api_key=api_key)
        response = client.messages.create(
            model="claude-3-haiku-20240307",
            max_tokens=500,
            system=system_prompt,
            messages=[{"role": "user", "content": user_prompt}],
        )

        content = response.content[0].text

        if verbose:
            console.print(f"         - API 응답 수신: {len(content)} chars")

        return _parse_ai_response(content, "anthropic", verbose)

    except ImportError:
        raise AIGenerationError("Anthropic library not installed. Run: pip install anthropic")
    except Exception as e:
        raise AIGenerationError(f"Anthropic API error: {e}")


def _parse_ai_response(response: str, provider: str, verbose: bool = False) -> Dict[str, str]:
    """Parse AI response and extract subject/body"""
    try:
        if verbose:
            from rich.console import Console

            console = Console()
            console.print("         - JSON 파싱 시도 중...")

        # Clean response - remove markdown formatting
        response = response.strip()
        if response.startswith("```json"):
            response = response[7:]
        if response.endswith("```"):
            response = response[:-3]
        response = response.strip()

        # Try to parse JSON
        try:
            data = json.loads(response)
            if "subject" in data and "body" in data:
                if verbose:
                    console.print("         - ✅ JSON 파싱 성공")
                return {"subject": data["subject"].strip(), "body": data["body"].strip()}
        except json.JSONDecodeError:
            if verbose:
                console.print("         - JSON 파싱 실패, 대체 방법 시도...")

        # Fallback: try to extract subject and body from text
        lines = response.split("\n")
        subject = ""
        body = ""

        for i, line in enumerate(lines):
            line = line.strip()
            if line.startswith('"subject"') or line.startswith("제목:"):
                subject = _extract_value(line)
            elif line.startswith('"body"') or line.startswith("본문:"):
                body = "\n".join(lines[i:])
                body = _extract_value(body)
                break

        if subject and body:
            return {"subject": subject, "body": body}

        # Last resort: use first line as subject, rest as body
        if lines:
            return {
                "subject": lines[0][:100],  # Limit subject length
                "body": "\n".join(lines[1:]) if len(lines) > 1 else lines[0],
            }

        raise AIGenerationError("Could not parse AI response")

    except Exception as e:
        raise AIGenerationError(f"Failed to parse {provider} response: {e}")


def _extract_value(text: str) -> str:
    """Extract value from JSON-like text"""
    # Remove JSON formatting
    text = text.replace('"subject":', "").replace('"body":', "")
    text = text.replace("제목:", "").replace("본문:", "")
    text = text.strip(' "{},:')
    return text
