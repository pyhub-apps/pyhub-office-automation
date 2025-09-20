"""
pyhub-office-automation 메인 CLI 명령어
AI 에이전트와의 연동을 위한 구조화된 출력 제공
"""

import json
import click
import sys
import importlib
import inspect
from pathlib import Path

from ..version import get_version_info, get_version


@click.group()
@click.version_option(version=get_version(), prog_name="oa")
@click.pass_context
def cli(ctx):
    """
    pyhub-office-automation: AI 에이전트를 위한 Office 자동화 도구

    Excel 및 HWP 문서 자동화를 위한 CLI 명령어들을 제공합니다.
    주로 Gemini CLI 등의 AI 에이전트와 함께 사용됩니다.
    """
    ctx.ensure_object(dict)


@cli.command()
@click.option('--format', 'output_format', default='json',
              type=click.Choice(['json', 'text']),
              help='출력 형식 선택')
def info(output_format):
    """패키지 정보 및 설치 상태 출력"""
    try:
        version_info = get_version_info()

        # 의존성 상태 확인
        dependencies = check_dependencies()

        info_data = {
            "package": "pyhub-office-automation",
            "version": version_info,
            "platform": sys.platform,
            "python_version": sys.version,
            "dependencies": dependencies,
            "status": "installed"
        }

        if output_format == 'json':
            click.echo(json.dumps(info_data, ensure_ascii=False, indent=2))
        else:
            click.echo(f"Package: {info_data['package']}")
            click.echo(f"Version: {info_data['version']['version']}")
            click.echo(f"Platform: {info_data['platform']}")
            click.echo(f"Python: {info_data['python_version']}")
            click.echo("Dependencies:")
            for dep, status in info_data['dependencies'].items():
                status_mark = "✓" if status['available'] else "✗"
                click.echo(f"  {status_mark} {dep}: {status['version'] or 'Not installed'}")

    except Exception as e:
        error_data = {
            "error": str(e),
            "command": "info",
            "version": get_version()
        }
        click.echo(json.dumps(error_data, ensure_ascii=False, indent=2), err=True)
        sys.exit(1)


@cli.command()
@click.option('--format', 'output_format', default='text',
              type=click.Choice(['json', 'text']),
              help='출력 형식 선택')
def install_guide(output_format):
    """설치 가이드 출력"""
    guide_steps = [
        {
            "step": 1,
            "title": "Python 설치",
            "description": "Python 3.13 이상을 설치하세요",
            "url": "https://www.python.org/downloads/",
            "command": None
        },
        {
            "step": 2,
            "title": "패키지 설치",
            "description": "pip를 사용하여 pyhub-office-automation을 설치하세요",
            "command": "pip install pyhub-office-automation"
        },
        {
            "step": 3,
            "title": "설치 확인",
            "description": "oa 명령어가 정상 동작하는지 확인하세요",
            "command": "oa info"
        },
        {
            "step": 4,
            "title": "Excel 사용 시 (선택사항)",
            "description": "Microsoft Excel이 설치되어 있어야 합니다",
            "note": "xlwings는 Excel이 설치된 환경에서 동작합니다"
        },
        {
            "step": 5,
            "title": "HWP 사용 시 (선택사항, Windows 전용)",
            "description": "한글(HWP) 프로그램이 설치되어 있어야 합니다",
            "note": "pyhwpx는 Windows COM을 통해 HWP와 연동됩니다"
        }
    ]

    guide_data = {
        "title": "pyhub-office-automation 설치 가이드",
        "version": get_version(),
        "platform_requirement": "Windows 10/11 (추천)",
        "python_requirement": "Python 3.13+",
        "steps": guide_steps
    }

    if output_format == 'json':
        click.echo(json.dumps(guide_data, ensure_ascii=False, indent=2))
    else:
        click.echo(f"=== {guide_data['title']} ===")
        click.echo(f"Version: {guide_data['version']}")
        click.echo(f"Platform: {guide_data['platform_requirement']}")
        click.echo(f"Python: {guide_data['python_requirement']}")
        click.echo()

        for step in guide_steps:
            click.echo(f"Step {step['step']}: {step['title']}")
            click.echo(f"  {step['description']}")
            if step.get('command'):
                click.echo(f"  Command: {step['command']}")
            if step.get('url'):
                click.echo(f"  URL: {step['url']}")
            if step.get('note'):
                click.echo(f"  Note: {step['note']}")
            click.echo()


def discover_excel_commands():
    """Excel 디렉토리에서 사용 가능한 명령어들을 동적으로 발견"""
    excel_dir = Path(__file__).parent.parent / "excel"
    commands = []

    if not excel_dir.exists():
        return commands

    for py_file in excel_dir.glob("*.py"):
        if py_file.name.startswith("__") or py_file.name in ["__init__.py"]:
            continue

        try:
            # 모듈 이름 생성 (파일명에서 .py 제거)
            module_name = py_file.stem
            module_path = f"pyhub_office_automation.excel.{module_name}"

            # 동적으로 모듈 import
            module = importlib.import_module(module_path)

            # click 명령어 찾기 (함수 또는 Command 객체)
            for name, obj in inspect.getmembers(module):
                if name.startswith('_'):  # private 멤버 제외
                    continue

                is_click_command = False

                # click.core.Command 타입인지 확인
                if hasattr(obj, '__class__') and 'click.core.Command' in str(type(obj)):
                    is_click_command = True

                # 또는 click 데코레이터가 적용된 함수인지 확인
                elif (inspect.isfunction(obj) and
                      hasattr(obj, '__click_params__')):
                    is_click_command = True

                if is_click_command:
                    # 명령어 이름을 파일명에서 가져오기 (snake_case를 kebab-case로)
                    command_name = module_name.replace('_', '-')

                    # 명령어 설명 추출
                    description = "설명 없음"

                    # Command 객체에서 help 가져오기
                    if hasattr(obj, 'help') and obj.help:
                        description = obj.help.strip()
                    # 또는 docstring에서 가져오기
                    elif hasattr(obj, '__doc__') and obj.__doc__:
                        # docstring의 첫 번째 줄만 사용
                        description = obj.__doc__.strip().split('\n')[0]
                    # 함수의 경우 callback에서 가져오기
                    elif hasattr(obj, 'callback') and obj.callback and obj.callback.__doc__:
                        description = obj.callback.__doc__.strip().split('\n')[0]

                    # 설명이 여러 줄인 경우 첫 번째 줄만 사용
                    if '\n' in description:
                        description = description.split('\n')[0].strip()

                    commands.append({
                        "name": command_name,
                        "description": description,
                        "module": module_path,
                        "function": name,
                        "status": "available"
                    })
                    break  # 첫 번째 click 명령어만 사용

        except Exception as e:
            # 모듈 import 실패 시 계속 진행
            commands.append({
                "name": py_file.stem.replace('_', '-'),
                "description": f"로드 실패: {str(e)}",
                "status": "error"
            })
            continue

    return commands


@cli.group()
def excel():
    """Excel 자동화 명령어들"""
    pass


def register_excel_commands():
    """발견된 Excel 명령어들을 동적으로 등록"""
    commands = discover_excel_commands()

    for cmd_info in commands:
        if cmd_info["status"] == "available":
            try:
                # 모듈 import 및 함수 가져오기
                module = importlib.import_module(cmd_info["module"])
                command_func = getattr(module, cmd_info["function"])

                # Excel 그룹에 명령어 추가
                excel.add_command(command_func, name=cmd_info["name"])

            except Exception as e:
                # 등록 실패 시 로그만 남기고 계속 진행
                pass


# Excel 명령어들을 동적으로 등록
register_excel_commands()


@excel.command(name='list')
@click.option('--format', 'output_format', default='json',
              type=click.Choice(['json', 'text']),
              help='출력 형식 선택')
def excel_list(output_format):
    """Excel 자동화 명령어 목록 출력"""
    # 실제 구현된 명령어들을 동적으로 스캔
    commands = discover_excel_commands()

    # 추가 정보 보강
    for cmd in commands:
        cmd["version"] = get_version()  # 패키지 버전 사용

    excel_data = {
        "category": "excel",
        "description": "Excel 자동화 명령어들 (xlwings 기반)",
        "commands": commands,
        "total_commands": len(commands),
        "available_commands": len([cmd for cmd in commands if cmd["status"] == "available"]),
        "package_version": get_version()
    }

    if output_format == 'json':
        click.echo(json.dumps(excel_data, ensure_ascii=False, indent=2))
    else:
        click.echo(f"=== Excel 자동화 명령어 목록 ===")
        click.echo(f"Total: {excel_data['total_commands']} commands")
        click.echo(f"Available: {excel_data['available_commands']} commands")
        click.echo()
        for cmd in commands:
            if cmd["status"] == "available":
                status_mark = "✅"
            elif cmd["status"] == "error":
                status_mark = "❌"
            else:
                status_mark = "○"

            click.echo(f"  {status_mark} oa excel {cmd['name']}")
            click.echo(f"     {cmd['description']}")
            if cmd["status"] == "available":
                click.echo(f"     버전: {cmd.get('version', 'unknown')}")
            click.echo()


@cli.group()
def hwp():
    """HWP 자동화 명령어들 (Windows 전용)"""
    pass


@hwp.command(name='list')
@click.option('--format', 'output_format', default='json',
              type=click.Choice(['json', 'text']),
              help='출력 형식 선택')
def hwp_list(output_format):
    """HWP 자동화 명령어 목록 출력"""
    # TODO: 실제 HWP 명령어들이 구현된 후 동적으로 스캔
    commands = [
        {
            "name": "open-hwp",
            "description": "HWP 파일 열기",
            "version": "1.0.0",
            "status": "planned"
        },
        {
            "name": "save-hwp",
            "description": "HWP 파일 저장",
            "version": "1.0.0",
            "status": "planned"
        }
    ]

    hwp_data = {
        "category": "hwp",
        "description": "HWP 자동화 명령어들 (pyhwpx 기반, Windows 전용)",
        "platform_requirement": "Windows + HWP 프로그램 설치 필요",
        "commands": commands,
        "total_commands": len(commands),
        "package_version": get_version()
    }

    if output_format == 'json':
        click.echo(json.dumps(hwp_data, ensure_ascii=False, indent=2))
    else:
        click.echo(f"=== HWP 자동화 명령어 목록 ===")
        click.echo(f"Platform: {hwp_data['platform_requirement']}")
        click.echo(f"Total: {hwp_data['total_commands']} commands")
        click.echo()
        for cmd in commands:
            status_mark = "✓" if cmd['status'] == 'available' else "○"
            click.echo(f"  {status_mark} oa hwp {cmd['name']}")
            click.echo(f"     {cmd['description']} (v{cmd['version']})")


@cli.command()
@click.argument('category', type=click.Choice(['excel', 'hwp']))
@click.argument('command_name')
@click.option('--format', 'output_format', default='text',
              type=click.Choice(['json', 'text']),
              help='출력 형식 선택')
def get_help(category, command_name, output_format):
    """특정 명령어의 도움말 조회"""
    # TODO: 실제 명령어들이 구현된 후 동적으로 도움말 수집
    help_data = {
        "category": category,
        "command": command_name,
        "help": f"oa {category} {command_name} 명령어 도움말 (구현 예정)",
        "usage": f"oa {category} {command_name} [OPTIONS]",
        "status": "planned",
        "version": get_version()
    }

    if output_format == 'json':
        click.echo(json.dumps(help_data, ensure_ascii=False, indent=2))
    else:
        click.echo(f"Command: oa {category} {command_name}")
        click.echo(f"Usage: {help_data['usage']}")
        click.echo(f"Status: {help_data['status']}")
        click.echo()
        click.echo(help_data['help'])


def check_dependencies():
    """의존성 패키지 설치 상태 확인"""
    dependencies = {}

    # xlwings 확인
    try:
        import xlwings
        dependencies['xlwings'] = {
            'available': True,
            'version': xlwings.__version__
        }
    except ImportError:
        dependencies['xlwings'] = {
            'available': False,
            'version': None
        }

    # pyhwpx 확인 (Windows 전용)
    try:
        import pyhwpx
        dependencies['pyhwpx'] = {
            'available': True,
            'version': getattr(pyhwpx, '__version__', 'unknown')
        }
    except ImportError:
        dependencies['pyhwpx'] = {
            'available': False,
            'version': None
        }

    # pandas 확인
    try:
        import pandas
        dependencies['pandas'] = {
            'available': True,
            'version': pandas.__version__
        }
    except ImportError:
        dependencies['pandas'] = {
            'available': False,
            'version': None
        }

    return dependencies


if __name__ == '__main__':
    cli()