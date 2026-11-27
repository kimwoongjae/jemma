"""UTC storage and Korea/Japan (UTC+9) display helpers."""

from datetime import datetime, timezone
from zoneinfo import ZoneInfo


JAPAN_TIMEZONE = ZoneInfo('Asia/Tokyo')


def utc_now_naive():
    """Return a naive UTC datetime for legacy timestamp columns."""
    return datetime.now(timezone.utc).replace(tzinfo=None)


def japan_now():
    return datetime.now(JAPAN_TIMEZONE)


def format_japan_datetime(value, date_format='%Y-%m-%d %H:%M'):
    """Format a timestamp as Japan time; naive stored values are interpreted as UTC."""
    if value is None:
        return ''
    if isinstance(value, datetime):
        parsed = value
    else:
        raw = str(value).strip()
        if not raw:
            return ''
        try:
            parsed = datetime.fromisoformat(raw.replace('Z', '+00:00'))
        except (TypeError, ValueError):
            return raw[:16]
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone.utc)
    return parsed.astimezone(JAPAN_TIMEZONE).strftime(date_format)


def japan_date_sql(expression, use_postgres=False):
    """SQL expression returning the UTC-stored timestamp's Japan calendar date."""
    if use_postgres:
        return (
            "TO_CHAR(CAST(NULLIF(CAST((" + expression + ") AS TEXT), '') AS TIMESTAMP) "
            "+ INTERVAL '9 hours', 'YYYY-MM-DD')"
        )
    return "SUBSTR(DATETIME(" + expression + ", '+9 hours'), 1, 10)"
