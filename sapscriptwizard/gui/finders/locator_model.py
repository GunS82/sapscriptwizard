"""Data structures describing GUI element semantics and locator hints."""
from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from typing import Any, Mapping, MutableMapping


class LocatorType(str, Enum):
    H_LABEL = "HLabel"
    V_LABEL = "VLabel"
    BY_TITLE = "ByTitle"
    BY_COL_ROW = "ByColRow"
    BY_ID = "ById"


@dataclass(slots=True)
class ElementInfo:
    element_id: str
    label: str | None
    value: Any
    element_type: str
    tech_name: str | None

    def to_dict(self) -> MutableMapping[str, Any]:
        payload: MutableMapping[str, Any] = {
            "id": self.element_id,
            "type": self.element_type,
            "label": self.label,
            "value": self.value,
        }
        if self.tech_name:
            payload["tech_name"] = self.tech_name
        return payload


@dataclass(slots=True)
class LocatorHint:
    locator_type: LocatorType
    query: str
    confidence: float
    metadata: Mapping[str, Any] | None = None

    def to_dict(self) -> MutableMapping[str, Any]:
        payload: MutableMapping[str, Any] = {
            "type": self.locator_type.value,
            "query": self.query,
            "confidence": self.confidence,
        }
        if self.metadata:
            payload["metadata"] = dict(self.metadata)
        return payload


__all__ = ["LocatorType", "ElementInfo", "LocatorHint"]
