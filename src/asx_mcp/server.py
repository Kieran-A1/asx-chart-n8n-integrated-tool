from __future__ import annotations

import asyncio
import os
import time
from pathlib import Path

import anyio
from mcp.server.fastmcp import FastMCP

from .pipeline import AsxReportError, normalize_asx_code, run_asx_report

mcp = FastMCP(
    "asx-report-agent",
    instructions=(
        "Capture a Yahoo Finance AU stock chart screenshot, place it in a Microsoft Word "
        "document, convert to PDF, and email it through Apple Mail."
    ),
)


_RECENT_RESULTS: dict[str, tuple[float, dict[str, str]]] = {}
_RECENT_ERRORS: dict[str, tuple[float, str]] = {}
_IN_FLIGHT: dict[str, asyncio.Task[dict[str, str]]] = {}
_IN_FLIGHT_LOCK = asyncio.Lock()


def _canonical_ticker(asx_code: str) -> str:
    return normalize_asx_code(asx_code) or "HUB"


def _canonical_recipient(recipient: str) -> str:
    value = (recipient or "").strip().lower()
    return value or "test@gmail.com"


def _request_key(asx_code: str, recipient: str, send_email: bool) -> str:
    return "|".join(
        [
            _canonical_ticker(asx_code),
            _canonical_recipient(recipient),
            str(bool(send_email)).lower(),
        ]
    )


def _dedupe_window_seconds() -> int:
    raw = os.getenv("MCP_DEDUPE_SECONDS", "90")
    try:
        return max(0, int(raw))
    except ValueError:
        return 90


def _error_dedupe_window_seconds() -> int:
    raw = os.getenv("MCP_ERROR_DEDUPE_SECONDS", "25")
    try:
        return max(0, int(raw))
    except ValueError:
        return 25


def _prune_cache(now: float, success_window: int, error_window: int) -> None:
    if success_window > 0:
        for key, (ts, _) in list(_RECENT_RESULTS.items()):
            if now - ts > success_window * 2:
                _RECENT_RESULTS.pop(key, None)
    else:
        _RECENT_RESULTS.clear()

    if error_window > 0:
        for key, (ts, _) in list(_RECENT_ERRORS.items()):
            if now - ts > error_window * 2:
                _RECENT_ERRORS.pop(key, None)
    else:
        _RECENT_ERRORS.clear()


async def _run_report(
    asx_code: str,
    recipient: str,
    output_dir: str,
    email_subject: str,
    email_body: str,
    send_email: bool,
) -> dict[str, str]:
    return await anyio.to_thread.run_sync(
        run_asx_report,
        asx_code or None,
        recipient,
        Path(output_dir),
        email_subject or None,
        email_body or None,
        send_email,
    )


@mcp.tool()
async def create_asx_report(
    asx_code: str = "",
    recipient: str = "test@gmail.com",
    output_dir: str = "output",
    email_subject: str = "",
    email_body: str = "",
    send_email: bool = True,
) -> dict[str, str]:
    """Create and optionally email a Yahoo Finance chart report.

    Args:
        asx_code: Ticker code like BHP, CBA, CSL. Leave empty for HUB. URL uses {CODE}.AX on Yahoo Finance AU.
        recipient: Email recipient address.
        output_dir: Local folder for image/docx/pdf artifacts. Do not pass placeholder paths like /path/to/local/directory/.
        email_subject: Optional custom email subject.
        email_body: Optional custom email body.
        send_email: False to generate files without sending email.
    """
    request_key = _request_key(
        asx_code=asx_code,
        recipient=recipient,
        send_email=send_email,
    )

    success_window = _dedupe_window_seconds()
    error_window = _error_dedupe_window_seconds()
    now = time.monotonic()

    _prune_cache(now=now, success_window=success_window, error_window=error_window)

    cached = _RECENT_RESULTS.get(request_key)
    if success_window > 0 and cached and now - cached[0] <= success_window:
        return {
            **cached[1],
            "deduplicated": "true",
            "dedupe_reason": "recent-cache",
            "dedupe_window_seconds": str(success_window),
        }

    cached_error = _RECENT_ERRORS.get(request_key)
    if error_window > 0 and cached_error and now - cached_error[0] <= error_window:
        raise AsxReportError(
            f"Skipped duplicate retry after recent failure: {cached_error[1]}"
        )

    owns_task = False
    async with _IN_FLIGHT_LOCK:
        task = _IN_FLIGHT.get(request_key)
        if task is None:
            task = asyncio.create_task(
                _run_report(
                    asx_code=asx_code,
                    recipient=recipient,
                    output_dir=output_dir,
                    email_subject=email_subject,
                    email_body=email_body,
                    send_email=send_email,
                )
            )
            _IN_FLIGHT[request_key] = task
            owns_task = True

    try:
        result = await task
    except Exception as exc:
        if owns_task:
            async with _IN_FLIGHT_LOCK:
                _IN_FLIGHT.pop(request_key, None)
            _RECENT_ERRORS[request_key] = (time.monotonic(), str(exc))
        raise AsxReportError(f"Failed to create Yahoo chart report: {exc}") from exc

    if owns_task:
        async with _IN_FLIGHT_LOCK:
            _IN_FLIGHT.pop(request_key, None)
        _RECENT_ERRORS.pop(request_key, None)
        _RECENT_RESULTS[request_key] = (time.monotonic(), dict(result))
        return {
            **result,
            "deduplicated": "false",
            "dedupe_reason": "none",
            "dedupe_window_seconds": str(success_window),
        }

    return {
        **result,
        "deduplicated": "true",
        "dedupe_reason": "in-flight",
        "dedupe_window_seconds": str(success_window),
    }


def main() -> None:
    transport = os.getenv("MCP_TRANSPORT", "sse").strip().lower()
    mount_path = os.getenv("MCP_MOUNT_PATH", "/")
    mcp.settings.host = os.getenv("MCP_HOST", "127.0.0.1")
    mcp.settings.port = int(os.getenv("MCP_PORT", "8001"))
    mcp.settings.sse_path = os.getenv("MCP_SSE_PATH", "/sse")
    mcp.settings.message_path = os.getenv("MCP_MESSAGE_PATH", "/messages/")
    mcp.settings.streamable_http_path = os.getenv("MCP_STREAMABLE_HTTP_PATH", "/mcp")

    if transport in {"streamable-http", "streamable_http", "http"}:
        transport = "streamable-http"

    mcp.run(transport=transport, mount_path=mount_path)


if __name__ == "__main__":
    main()
