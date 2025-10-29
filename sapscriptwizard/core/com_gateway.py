"""Shared COM interaction utilities with retry and timeout policies."""
from __future__ import annotations

from dataclasses import dataclass
import logging
import time
from typing import Any, Callable, Iterable, Protocol, Sequence, TypeVar

from .errors import SapError, SapErrorCode


_LOGGER = logging.getLogger(__name__)

T = TypeVar("T")


class RetryableError(Protocol):
    args: Sequence[Any]


@dataclass(slots=True)
class RetryPolicy:
    """Configuration for retry attempts."""

    max_attempts: int = 3
    base_delay: float = 0.2
    multiplier: float = 2.0


class ComGateway:
    """Executes COM calls with centralized retry and logging."""

    def __init__(
        self,
        retry_policy: RetryPolicy | None = None,
        retry_exceptions: Iterable[type[Exception]] | None = None,
        logger: logging.Logger | None = None,
    ) -> None:
        self.retry_policy = retry_policy or RetryPolicy()
        self.retry_exceptions = tuple(retry_exceptions or ())
        self.logger = logger or _LOGGER

    def call(
        self,
        func: Callable[..., T],
        *args: Any,
        error_code: SapErrorCode = SapErrorCode.COM_FAIL,
        **kwargs: Any,
    ) -> T:
        """Call ``func`` with retry semantics."""

        attempts = 0
        delay = self.retry_policy.base_delay
        last_error: Exception | None = None

        while attempts < self.retry_policy.max_attempts:
            try:
                result = func(*args, **kwargs)
                self.logger.debug("COM call %s succeeded", getattr(func, "__name__", repr(func)))
                return result
            except Exception as exc:  # noqa: BLE001 - we need to inspect exception dynamically
                last_error = exc
                attempts += 1
                is_retryable = isinstance(exc, self.retry_exceptions)
                self.logger.warning(
                    "COM call %s failed (%s/%s): %s",
                    getattr(func, "__name__", repr(func)),
                    attempts,
                    self.retry_policy.max_attempts,
                    exc,
                )
                if not is_retryable or attempts >= self.retry_policy.max_attempts:
                    break
                time.sleep(delay)
                delay *= self.retry_policy.multiplier

        raise SapError(
            error_code,
            "COM call failed after retries",
            {
                "function": getattr(func, "__name__", repr(func)),
                "args": args,
                "kwargs": kwargs,
                "attempts": attempts,
                "last_error": repr(last_error),
            },
        )

    def wrap_attr(self, obj: Any, attr: str) -> Callable[..., Any]:
        """Return a callable that executes ``getattr(obj, attr)`` via :meth:`call`."""

        def _wrapper(*args: Any, **kwargs: Any) -> Any:
            member = getattr(obj, attr)
            return self.call(member, *args, **kwargs)

        return _wrapper
