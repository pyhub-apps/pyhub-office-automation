"""
Email automation module for pyhub-office-automation
AI-powered email generation and sending functionality
"""

from .email_accounts import accounts_app, delete_email_account, list_email_accounts
from .email_send import email_send

__all__ = ["email_send", "accounts_app", "list_email_accounts", "delete_email_account"]
