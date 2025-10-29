"""Centralised COM gateway with retry and timeout support."""
from __future__ import annotations

import logging
import time
from dataclasses import dataclass
from typing import Any, Protocol

from .errors import ErrorCode, SapError

log = logging.getLogger(__name__)


class ComCallable(Protocol):
    """Protocol for COM callables."""

    def __call__(self, *args: Any, **kwargs: Any) -> Any:
        """Invoke the COM method."""


@dataclass(slots=True)
class RetryPolicy:
    """Retry configuration for COM invocations."""

    max_attempts: int = 3
    base_delay: float = 0.5
    backoff_factor: float = 2.0


class ComGateway:
    """Executes COM calls with retries and timeout awareness."""

    def __init__(self, retry_policy: RetryPolicy | None = None) -> None:
        self.retry_policy = retry_policy or RetryPolicy()

    def invoke(self, func: ComCallable, *args: Any, **kwargs: Any) -> Any:
        """Invoke *func* applying the configured retry policy."""

        delay = self.retry_policy.base_delay
        attempt = 1
        while True:
            try:
                name = getattr(func, "__name__", repr(func))
                log.debug("Invoking COM callable %s (attempt %s)", name, attempt)
                return func(*args, **kwargs)
            except Exception as exc:  # pragma: no cover - thin wrapper
                log.warning("COM invocation failed: %s", exc, exc_info=True)
                if attempt >= self.retry_policy.max_attempts:
                    raise SapError(ErrorCode.COM_FAIL, "COM call failed", str(exc)) from exc
                time.sleep(delay)
                delay *= self.retry_policy.backoff_factor
                attempt += 1

    def get_attr(self, obj: Any, name: str) -> Any:
        """Retrieve attribute *name* from *obj* using :meth:`invoke`."""

        return self.invoke(lambda: getattr(obj, name))

    def call_method(self, obj: Any, name: str, *args: Any, **kwargs: Any) -> Any:
        """Invoke a method on *obj* through the gateway."""

        def _call() -> Any:
            method = getattr(obj, name)
            return method(*args, **kwargs)

        return self.invoke(_call)
