"""Asynchronous Telegram delivery for tenant progress updates."""

import logging
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime

import requests

from database import get_connection


logger = logging.getLogger(__name__)
_telegram_executor = ThreadPoolExecutor(max_workers=2, thread_name_prefix='telegram-progress')


def send_telegram_message(bot_token, chat_id, text):
    bot_token = str(bot_token or '').strip()
    chat_id = str(chat_id or '').strip()
    if not bot_token or not chat_id:
        return False, '봇 토큰과 채팅 ID를 확인해주세요.'

    try:
        response = requests.post(
            f'https://api.telegram.org/bot{bot_token}/sendMessage',
            json={
                'chat_id': chat_id,
                'text': str(text or '')[:4000],
                'disable_web_page_preview': True
            },
            timeout=(5, 10)
        )
        try:
            payload = response.json()
        except ValueError:
            payload = {}
        if response.ok and payload.get('ok'):
            return True, '메시지를 전송했습니다.'
        description = str(payload.get('description') or f'Telegram HTTP {response.status_code}')
        return False, description[:240]
    except Exception as error:
        logger.warning('Telegram message delivery failed (%s)', type(error).__name__)
        return False, '텔레그램 서버에 연결하지 못했습니다.'


def _send_for_tenant(tenant_id, message):
    conn = None
    try:
        conn = get_connection()
        cur = conn.cursor()
        cur.execute('''SELECT bot_token, chat_id, enabled FROM telegram_settings
            WHERE tenant_id = ? LIMIT 1''', (int(tenant_id),))
        settings = cur.fetchone()
        if not settings or not int(settings['enabled'] or 0):
            return
        success, result = send_telegram_message(settings['bot_token'], settings['chat_id'], message)
        if not success:
            logger.warning('Telegram delivery failed for tenant %s: %s', tenant_id, result)
    except Exception:
        logger.exception('Could not deliver progress update to Telegram for tenant %s', tenant_id)
    finally:
        if conn:
            conn.close()


def queue_telegram_message(tenant_id, message):
    """Queue one progress notification without delaying the business transaction."""
    if not tenant_id or not str(message or '').strip():
        return
    try:
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        formatted = f'진행사항\n{timestamp}\n{str(message).strip()}'
        _telegram_executor.submit(_send_for_tenant, int(tenant_id), formatted)
    except RuntimeError:
        logger.exception('Telegram notification queue is unavailable')
