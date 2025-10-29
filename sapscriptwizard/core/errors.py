"""Error definitions for the SAP Script Wizard core layer."""
from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from typing import Any, Mapping, MutableMapping, Optional


class SapErrorCode(str, Enum):
    """Canonical error codes for all SAP GUI interactions."""

    NOT_FOUND = "E_NOT_FOUND"
    DISABLED = "E_DISABLED"
    TIMEOUT = "E_TIMEOUT"
    COM_FAIL = "E_COM_FAIL"
    UNSUPPORTED = "E_UNSUPPORTED"
    AMBIGUOUS = "E_AMBIGUOUS"


@dataclass(slots=True)
class SapError(Exception):
    """Rich error object that preserves structured diagnostic information."""

    code: SapErrorCode
    message: str
    details: Optional[Mapping[str, Any]] = None

    def to_dict(self) -> MutableMapping[str, Any]:
        """Serialize the error to a JSON friendly structure."""

        payload: MutableMapping[str, Any] = {
            "code": self.code.value,
            "message": self.message,
        }
        if self.details:
            payload["details"] = dict(self.details)
        return payload

    def __str__(self) -> str:  # pragma: no cover - trivial wrapper
        if self.details:
            return f"{self.code.value}: {self.message} ({self.details})"
        return f"{self.code.value}: {self.message}"
