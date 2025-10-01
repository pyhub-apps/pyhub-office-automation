"""
Interactive shell module for pyhub-office-automation
Provides stateful REPL interfaces for Excel and PowerPoint automation
"""

from pyhub_office_automation.shell.excel_shell import excel_shell
from pyhub_office_automation.shell.ppt_shell import ppt_shell
from pyhub_office_automation.shell.unified_shell import unified_shell

__all__ = ["excel_shell", "ppt_shell", "unified_shell"]
