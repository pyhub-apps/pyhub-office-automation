"""
AI 에이전트별 맞춤형 설정 파일 자동 생성 CLI 모듈
GitHub Issue #56 구현
"""

import json
import os
import re
import sys
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import typer
from rich.console import Console
from rich.panel import Panel
from rich.table import Table

from pyhub_office_automation.utils.python_detector import PythonDetector, get_best_python
from pyhub_office_automation.utils.resource_loader import get_resource_path, load_resource_text
from pyhub_office_automation.version import get_version

# Typer 앱 생성
ai_setup_app = typer.Typer(help="AI 에이전트별 맞춤형 설정 파일 자동 생성", no_args_is_help=True)

# Rich 콘솔
console = Console()

# 지원하는 AI 에이전트 정보
AI_AGENTS = {
    "codex": {
        "name": "Codex CLI",
        "filename": "AGENTS.md",
        "description": "GitHub Codex CLI용 지침 파일",
        "template": "codex_template.md"
    },
    "gemini": {
        "name": "Gemini CLI",
        "filename": "GEMINI.md",
        "description": "Google Gemini CLI용 지침 파일",
        "template": "gemini_template.md"
    },
    "claude": {
        "name": "Claude Code",
        "filename": "CLAUDE.md",
        "description": "Anthropic Claude Code용 지침 파일",
        "template": "claude_template.md"
    }
}


class AISetupManager:
    """AI 설정 파일 생성 및 관리 클래스"""

    def __init__(self, output_dir: Path = None):
        self.output_dir = output_dir or Path.cwd()
        self.python_detector = PythonDetector()
        self.timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.package_version = get_version()

    def generate_ai_setup(self, agent_type: str, detect_python: bool = True,
                         force: bool = False, dry_run: bool = False) -> Dict:
        """AI 에이전트별 설정 파일 생성"""
        if agent_type not in AI_AGENTS:
            raise ValueError(f"지원하지 않는 AI 에이전트: {agent_type}")

        agent_info = AI_AGENTS[agent_type]
        output_file = self.output_dir / agent_info["filename"]

        # 콘텐츠 생성
        content = self._build_content(agent_type, detect_python)

        # 기존 파일 처리
        if output_file.exists() and not force:
            content = self._merge_with_existing(output_file, content)

        result = {
            "agent_type": agent_type,
            "agent_name": agent_info["name"],
            "output_file": str(output_file),
            "content_length": len(content),
            "timestamp": self.timestamp,
            "sections": self._get_content_sections(content)
        }

        if not dry_run:
            # 파일 쓰기
            try:
                output_file.write_text(content, encoding="utf-8")
                result["status"] = "created" if not output_file.exists() else "updated"
                result["success"] = True
            except Exception as e:
                result["status"] = "error"
                result["error"] = str(e)
                result["success"] = False
        else:
            result["status"] = "dry_run"
            result["success"] = True
            result["preview"] = content[:500] + "..." if len(content) > 500 else content

        return result

    def generate_all_setups(self, detect_python: bool = True,
                           force: bool = False, dry_run: bool = False) -> List[Dict]:
        """모든 AI 에이전트 설정 파일 생성"""
        results = []
        for agent_type in AI_AGENTS.keys():
            try:
                result = self.generate_ai_setup(agent_type, detect_python, force, dry_run)
                results.append(result)
            except Exception as e:
                results.append({
                    "agent_type": agent_type,
                    "agent_name": AI_AGENTS[agent_type]["name"],
                    "status": "error",
                    "error": str(e),
                    "success": False
                })
        return results

    def get_status(self) -> Dict:
        """현재 AI 설정 파일 상태 확인"""
        status = {
            "timestamp": self.timestamp,
            "output_directory": str(self.output_dir),
            "python_detection": {},
            "ai_files": {}
        }

        # Python 감지 상태
        python_info = get_best_python()
        if python_info:
            status["python_detection"] = {
                "found": True,
                "path": python_info.path,
                "version": python_info.version,
                "is_recommended": python_info.is_recommended
            }
        else:
            status["python_detection"] = {"found": False}

        # AI 파일 상태
        for agent_type, agent_info in AI_AGENTS.items():
            file_path = self.output_dir / agent_info["filename"]
            if file_path.exists():
                stat = file_path.stat()
                status["ai_files"][agent_type] = {
                    "exists": True,
                    "filename": agent_info["filename"],
                    "size": stat.st_size,
                    "last_modified": datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
                    "has_ai_context": self._has_ai_context(file_path)
                }
            else:
                status["ai_files"][agent_type] = {
                    "exists": False,
                    "filename": agent_info["filename"]
                }

        return status

    def _build_content(self, agent_type: str, detect_python: bool) -> str:
        """AI 에이전트별 콘텐츠 생성"""
        sections = []

        # 기본 템플릿 로드
        base_content = self._load_template("base_template.md")
        if base_content:
            sections.append(base_content)

        # Python 환경 정보 추가
        if detect_python:
            python_content = self._build_python_section()
            if python_content:
                sections.append(python_content)

        # 차트 가이드 추가
        chart_content = self._load_template("chart_template.md")
        if chart_content:
            sections.append(chart_content)

        # AI별 특화 템플릿 추가
        ai_template = AI_AGENTS[agent_type]["template"]
        ai_content = self._load_template(ai_template)
        if ai_content:
            sections.append(ai_content)

        # 메타데이터 추가
        metadata = self._build_metadata(agent_type)
        sections.append(metadata)

        return "\n\n".join(sections)

    def _build_python_section(self) -> Optional[str]:
        """Python 환경 섹션 생성"""
        python_info = get_best_python()
        if not python_info:
            return None

        template = self._load_template("python_template.md")
        if not template:
            return None

        # Python 경로 치환
        return template.format(python_path=python_info.path)

    def _build_metadata(self, agent_type: str) -> str:
        """메타데이터 섹션 생성"""
        agent_info = AI_AGENTS[agent_type]
        return f"""
---

## 설정 파일 정보

- **생성 대상**: {agent_info['name']}
- **생성 시간**: {self.timestamp}
- **패키지 버전**: {self.package_version}
- **Python 탐지**: {"활성화" if get_best_python() else "비활성화"}

이 파일은 `oa ai-setup {agent_type}` 명령으로 생성되었습니다.
"""

    def _load_template(self, template_name: str) -> Optional[str]:
        """템플릿 파일 로드"""
        try:
            template_path = get_resource_path(f"ai_templates/{template_name}")
            if template_path and template_path.exists():
                return template_path.read_text(encoding="utf-8")
        except Exception as e:
            console.print(f"[yellow]Warning: Failed to load template {template_name}: {e}[/yellow]")
        return None

    def _merge_with_existing(self, file_path: Path, new_content: str) -> str:
        """기존 파일과 새 콘텐츠 병합"""
        try:
            existing_content = file_path.read_text(encoding="utf-8")

            # "# Code Assistant Context" 섹션이 있는지 확인
            context_pattern = r"^# Code Assistant Context.*?(?=^#|\Z)"
            context_match = re.search(context_pattern, existing_content, re.MULTILINE | re.DOTALL)

            if context_match:
                # 기존 섹션 교체
                updated_content = re.sub(context_pattern, new_content.strip(),
                                       existing_content, flags=re.MULTILINE | re.DOTALL)
                return updated_content
            else:
                # 파일 끝에 추가
                return existing_content.rstrip() + "\n\n" + new_content

        except Exception as e:
            console.print(f"[yellow]Warning: Failed to merge with existing file: {e}[/yellow]")
            return new_content

    def _has_ai_context(self, file_path: Path) -> bool:
        """파일에 AI 컨텍스트 섹션이 있는지 확인"""
        try:
            content = file_path.read_text(encoding="utf-8")
            return "# Code Assistant Context" in content
        except Exception:
            return False

    def _get_content_sections(self, content: str) -> List[str]:
        """콘텐츠에서 섹션 목록 추출"""
        sections = []
        lines = content.split("\n")
        for line in lines:
            if line.startswith("## ") and not line.startswith("## 설정 파일 정보"):
                sections.append(line[3:].strip())  # "## " 제거
        return sections


@ai_setup_app.command()
def setup(
    agent_type: str = typer.Argument(..., help="AI 에이전트 타입 (codex, gemini, claude, all)"),
    detect_python: bool = typer.Option(True, "--detect-python/--no-detect-python",
                                     help="Python 환경 자동 탐지 여부"),
    force: bool = typer.Option(False, "--force", help="기존 파일 덮어쓰기"),
    dry_run: bool = typer.Option(False, "--dry-run", help="실제 파일 생성 없이 미리보기"),
    output_dir: str = typer.Option(".", "--output-dir", help="출력 디렉토리 경로"),
    output_format: str = typer.Option("text", "--format", help="출력 형식 (text, json)")
):
    """AI 에이전트별 맞춤형 설정 파일 생성"""
    try:
        output_path = Path(output_dir).resolve()
        if not output_path.exists():
            output_path.mkdir(parents=True, exist_ok=True)

        manager = AISetupManager(output_path)

        if agent_type.lower() == "all":
            results = manager.generate_all_setups(detect_python, force, dry_run)
            _display_results(results, output_format, dry_run)
        else:
            if agent_type.lower() not in AI_AGENTS:
                available = ", ".join(AI_AGENTS.keys())
                typer.echo(f"Error: 지원하지 않는 AI 에이전트 '{agent_type}'. 사용 가능: {available}, all")
                raise typer.Exit(1)

            result = manager.generate_ai_setup(agent_type.lower(), detect_python, force, dry_run)
            _display_results([result], output_format, dry_run)

    except Exception as e:
        typer.echo(f"Error: {e}", err=True)
        raise typer.Exit(1)


@ai_setup_app.command()
def status(
    output_dir: str = typer.Option(".", "--output-dir", help="확인할 디렉토리 경로"),
    output_format: str = typer.Option("text", "--format", help="출력 형식 (text, json)")
):
    """현재 AI 설정 파일 상태 확인"""
    try:
        output_path = Path(output_dir).resolve()
        manager = AISetupManager(output_path)
        status_info = manager.get_status()

        if output_format == "json":
            typer.echo(json.dumps(status_info, ensure_ascii=False, indent=2))
        else:
            _display_status(status_info)

    except Exception as e:
        typer.echo(f"Error: {e}", err=True)
        raise typer.Exit(1)


def _display_results(results: List[Dict], output_format: str, dry_run: bool):
    """결과 출력"""
    if output_format == "json":
        output_data = {
            "results": results,
            "summary": {
                "total": len(results),
                "successful": sum(1 for r in results if r.get("success", False)),
                "failed": sum(1 for r in results if not r.get("success", False)),
                "dry_run": dry_run
            }
        }
        typer.echo(json.dumps(output_data, ensure_ascii=False, indent=2))
    else:
        _display_results_table(results, dry_run)


def _display_results_table(results: List[Dict], dry_run: bool):
    """결과를 테이블로 출력"""
    title = "🧪 AI 설정 파일 생성 미리보기" if dry_run else "✨ AI 설정 파일 생성 결과"

    table = Table(title=title)
    table.add_column("AI 에이전트", style="cyan")
    table.add_column("파일명", style="green")
    table.add_column("상태", style="bold")
    table.add_column("섹션 수", justify="right")

    for result in results:
        if result.get("success", False):
            status = "🔍 미리보기" if dry_run else f"✅ {result['status']}"
            sections_count = str(len(result.get("sections", [])))
        else:
            status = f"❌ {result.get('error', '실패')}"
            sections_count = "-"

        table.add_row(
            result.get("agent_name", result["agent_type"]),
            result.get("output_file", "").split("/")[-1] if result.get("output_file") else "-",
            status,
            sections_count
        )

    console.print(table)

    # 성공한 결과에 대한 추가 정보
    successful_results = [r for r in results if r.get("success", False)]
    if successful_results and not dry_run:
        console.print("\n📝 [bold green]생성된 파일 상세 정보:[/bold green]")
        for result in successful_results:
            if result.get("sections"):
                sections_text = ", ".join(result["sections"])
                console.print(f"  • {result['output_file']}: {sections_text}")


def _display_status(status_info: Dict):
    """상태 정보 출력"""
    console.print(Panel.fit("🔍 AI 설정 파일 상태 확인", style="bold blue"))

    # Python 환경 정보
    python_info = status_info["python_detection"]
    if python_info["found"]:
        python_text = f"✅ Python {python_info['version']} at {python_info['path']}"
        if python_info["is_recommended"]:
            python_text += " (권장 버전)"
    else:
        python_text = "❌ Python 환경을 찾을 수 없습니다"

    console.print(f"\n🐍 **Python 환경**: {python_text}")
    console.print(f"📁 **확인 디렉토리**: {status_info['output_directory']}")

    # AI 파일 상태 테이블
    table = Table(title="AI 설정 파일 상태")
    table.add_column("AI 에이전트", style="cyan")
    table.add_column("파일명", style="green")
    table.add_column("존재 여부", style="bold")
    table.add_column("AI 컨텍스트", style="yellow")
    table.add_column("최종 수정", style="dim")

    for agent_type, file_info in status_info["ai_files"].items():
        agent_name = AI_AGENTS[agent_type]["name"]
        filename = file_info["filename"]

        if file_info["exists"]:
            exists_status = "✅ 존재"
            context_status = "✅ 있음" if file_info["has_ai_context"] else "❌ 없음"
            last_modified = file_info["last_modified"]
        else:
            exists_status = "❌ 없음"
            context_status = "-"
            last_modified = "-"

        table.add_row(agent_name, filename, exists_status, context_status, last_modified)

    console.print(f"\n")
    console.print(table)

    # 권장 액션
    missing_files = [agent for agent, info in status_info["ai_files"].items() if not info["exists"]]
    files_without_context = [agent for agent, info in status_info["ai_files"].items()
                           if info["exists"] and not info["has_ai_context"]]

    if missing_files or files_without_context:
        console.print("\n💡 [bold yellow]권장 액션:[/bold yellow]")
        if missing_files:
            agents_list = ", ".join(missing_files)
            console.print(f"  • 누락된 파일 생성: [green]oa ai-setup {agents_list}[/green]")
        if files_without_context:
            agents_list = ", ".join(files_without_context)
            console.print(f"  • AI 컨텍스트 추가: [green]oa ai-setup {agents_list}[/green]")
        console.print(f"  • 모든 파일 생성: [green]oa ai-setup all[/green]")


if __name__ == "__main__":
    ai_setup_app()