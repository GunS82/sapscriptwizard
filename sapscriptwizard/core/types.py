"""Typed primitives shared across the core package."""
from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any


class LocatorHintType(str, Enum):
    H_LABEL = "HLabel"
    V_LABEL = "VLabel"
    BY_TITLE = "ByTitle"
    BY_COL_ROW = "ByColRow"


@dataclass(slots=True)
class ElementInfo:
    """Rich description of a GUI element."""

    element_id: str
    type: str
    tech_name: str | None = None
    label: str | None = None
    value: str | None = None
    enabled: bool | None = None
    visible: bool | None = None
    children: list[ElementInfo] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        """Serialise the element into a JSON-ready dictionary."""

        return {
            "id": self.element_id,
            "type": self.type,
            "tech_name": self.tech_name,
            "label": self.label,
            "value": self.value,
            "enabled": self.enabled,
            "visible": self.visible,
            "children": [child.to_dict() for child in self.children],
        }


@dataclass(slots=True)
class LocatorHint:
    """Suggestion for a more resilient locator."""

    hint_type: LocatorHintType
    data: dict[str, Any]

    def describe(self) -> str:
        """Human readable description."""

        match self.hint_type:
            case LocatorHintType.H_LABEL:
                return f"Horizontal label match: {self.data.get('label')}"
            case LocatorHintType.V_LABEL:
                return f"Vertical label match: {self.data.get('label')}"
            case LocatorHintType.BY_TITLE:
                return f"Element by title: {self.data.get('title')}"
            case LocatorHintType.BY_COL_ROW:
                return f"Table cell {self.data.get('row')}:{self.data.get('column')}"
        return "Unknown locator"
