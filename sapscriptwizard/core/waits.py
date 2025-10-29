"""Waiting helpers for synchronisation with SAP GUI."""
from __future__ import annotations

import time
from typing import Callable

from .errors import SapError, SapErrorCode


def wait_visible(window: "Window", element_id: str, timeout: float = 10.0, poll_interval: float = 0.2) -> None:
    from .window import Window

    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        try:
            element = window.find_by_id(element_id)
        except SapError as exc:
            if exc.code is SapErrorCode.NOT_FOUND:
                time.sleep(poll_interval)
                continue
            raise
        visible = getattr(element, "Visible", getattr(element, "visible", True))
        if visible:
            return
        time.sleep(poll_interval)
    raise SapError(SapErrorCode.TIMEOUT, f"Element {element_id} not visible after {timeout}s")


def wait_ready(window: "Window", timeout: float = 10.0, poll_interval: float = 0.2) -> None:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        busy = getattr(window._get_window(), "Busy", False)
        if not busy:
            return
        time.sleep(poll_interval)
    raise SapError(SapErrorCode.TIMEOUT, "Window not ready in time")


def wait_status(
    window: "Window",
    *,
    text: str | None = None,
    code: str | None = None,
    timeout: float = 10.0,
    poll_interval: float = 0.2,
) -> None:
    if text is None and code is None:
        raise ValueError("text or code must be provided")
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        status_bar = getattr(window._get_window(), "StatusBar", None)
        if status_bar is None:
            break
        current_text = str(getattr(status_bar, "Text", ""))
        current_code = str(getattr(status_bar, "MessageType", ""))
        if (text is None or current_text == text) and (code is None or current_code == code):
            return
        time.sleep(poll_interval)
    raise SapError(
        SapErrorCode.TIMEOUT,
        "Status did not reach desired state",
        {"text": text, "code": code},
    )


__all__ = ["wait_visible", "wait_ready", "wait_status"]
