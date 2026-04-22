"""Telegram notifier for margin-settlement monthly runs.

Reads credentials from env vars ``TELEGRAM_BOT_TOKEN`` / ``TELEGRAM_CHAT_ID``.
Call :func:`load_dotenv` once at startup to populate env from a ``.env`` file.

Design
------
Network failures are silenced (notification is best-effort; the tool's
primary job of writing settlement files must not be blocked by a Telegram
outage). The caller can inspect the return value to decide whether to log a
warning.
"""
from __future__ import annotations

import json
import os
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path


def load_dotenv(path: str | Path = ".env", *, override: bool = False) -> None:
    """Populate ``os.environ`` from a simple KEY=VALUE file.

    By default shell-level env vars take precedence. Pass ``override=True``
    to force the file values to win — useful when a user's shell has stale
    credentials from another tool.
    """
    p = Path(path)
    if not p.exists():
        return
    for raw in p.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        if "=" not in line:
            continue
        key, _, value = line.partition("=")
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if not key:
            continue
        if override or key not in os.environ:
            os.environ[key] = value


def send_telegram(
    message: str,
    *,
    token: str | None = None,
    chat_id: str | None = None,
    timeout: float = 10.0,
) -> bool:
    """Send a plain-text message to Telegram. Returns True on success.

    Falls back to env vars if ``token`` / ``chat_id`` omitted. Returns False
    silently when credentials are missing or the network call fails.
    """
    token = token or os.environ.get("TELEGRAM_BOT_TOKEN")
    chat_id = chat_id or os.environ.get("TELEGRAM_CHAT_ID")
    if not token or not chat_id:
        return False

    url = f"https://api.telegram.org/bot{token}/sendMessage"
    payload = urllib.parse.urlencode({
        "chat_id": chat_id,
        "text": message,
    }).encode("utf-8")

    try:
        req = urllib.request.Request(url, data=payload)
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            body = json.loads(resp.read().decode("utf-8"))
            return bool(body.get("ok"))
    except (urllib.error.URLError, TimeoutError, json.JSONDecodeError):
        return False


def format_run_summary(month: str, results: list[tuple[str, bool, str]]) -> str:
    """Format a Telegram-friendly summary of a monthly generation run.

    ``results`` is a list of ``(service, success, info)`` where ``info`` is
    the output path on success or an error message on failure.
    """
    ok_count = sum(1 for _, ok, _ in results if ok)
    total = len(results)

    lines = [f"【月次清算書生成】{month}"]
    lines.append(f"成功 {ok_count}/{total}")
    if total > 0:
        lines.append("")
        for service, ok, info in results:
            if ok:
                filename = info.replace("\\", "/").rsplit("/", 1)[-1]
                lines.append(f"✅ {service}: {filename}")
            else:
                lines.append(f"❌ {service}: {info}")
    return "\n".join(lines)
