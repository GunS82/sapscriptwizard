from __future__ import annotations

import time
from collections.abc import Callable

from .errors import ErrorCode, SapError

DEFAULT_TIMEOUT = 30.0
DEFAULT_INTERVAL = 0.5


def _wait_until(predicate: Callable[[], bool], timeout: float, interval: float) -> None:
    end_time = time.monotonic() + timeout
    while True:
        if predicate():
            return
        if time.monotonic() > end_time:
            raise SapError(ErrorCode.TIMEOUT, "Wait condition timed out")
        time.sleep(interval)


def wait_visible(
    getter: Callable[[], bool],
    timeout: float = DEFAULT_TIMEOUT,
    interval: float = DEFAULT_INTERVAL,
) -> None:
    """Wait until GUI element becomes visible."""

    _wait_until(getter, timeout, interval)


def wait_ready(
    status_getter: Callable[[], bool],
    timeout: float = DEFAULT_TIMEOUT,
    interval: float = DEFAULT_INTERVAL,
) -> None:
    """Wait until the main window reports ready state."""

    _wait_until(status_getter, timeout, interval)


def wait_status(
    checker: Callable[[str], bool],
    status_provider: Callable[[], str],
    *,
    timeout: float = DEFAULT_TIMEOUT,
    interval: float = DEFAULT_INTERVAL,
) -> None:
    """Wait until :func:`checker` returns True for the status text."""

    def _predicate() -> bool:
        return checker(status_provider())

    _wait_until(_predicate, timeout, interval)
