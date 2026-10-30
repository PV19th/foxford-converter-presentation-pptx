# -*- coding: utf-8 -*-
"""Все настройки сайта — из переменных окружения (файл .env)."""
import os, secrets
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA = Path(os.environ.get('DATA_DIR', ROOT / 'data'))
JOBS, LOGS = DATA / 'jobs', DATA / 'logs'

APP_PASSWORD = os.environ.get('APP_PASSWORD', '')
# без SECRET_KEY вход сбрасывается при каждом перезапуске сайта
SECRET_KEY = os.environ.get('SECRET_KEY') or secrets.token_hex(32)
JOB_TTL_HOURS = float(os.environ.get('JOB_TTL_HOURS') or 24)
MAX_UPLOAD_MB = int(os.environ.get('MAX_UPLOAD_MB') or 100)
SUPPORT_CONTACT = os.environ.get('SUPPORT_CONTACT', '')

# ИИ-доработка (этап 7): без ключа кнопки «Доработать с ИИ» на сайте нет
AI_ENABLED = bool(os.environ.get('ANTHROPIC_API_KEY'))
AI_DAILY_LIMIT_USD = float(os.environ.get('AI_DAILY_LIMIT_USD') or 20)
