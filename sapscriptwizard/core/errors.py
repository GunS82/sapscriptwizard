"""Error types for sapscriptwizard core layer."""
from __future__ import annotations

from dataclasses import dataclass
from enum import Enum


class ErrorCode(str, Enum):
    """Canonical error codes raised by the core layer."""

    NOT_FOUND = "E_NOT_FOUND"
    DISABLED = "E_DISABLED"
    TIMEOUT = "E_TIMEOUT"
    COM_FAIL = "E_COM_FAIL"
    UNSUPPORTED = "E_UNSUPPORTED"
    AMBIGUOUS = "E_AMBIGUOUS"


@dataclass(slots=True)
class SapError(Exception):
    """Structured SAP GUI scripting error."""

    code: ErrorCode
    message: str
    details: str | None = None

    def __str__(self) -> str:  # pragma: no cover - trivial
        detail = f" Details: {self.details}" if self.details else ""
        return f"[{self.code}] {self.message}{detail}"
