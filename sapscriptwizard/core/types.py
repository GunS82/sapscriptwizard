"""Shared type definitions for the core package."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Iterable, Mapping, MutableMapping, Sequence


@dataclass(slots=True)
class ElementState:
    """Snapshot of a GUI element state returned from :func:`Window.dump_gui_structure`."""

    element_id: str
    element_type: str
    tech_name: str | None
    label: str | None
    value: Any
    enabled: bool
    visible: bool
    children: Sequence["ElementState"] = field(default_factory=tuple)

    def to_dict(self) -> MutableMapping[str, Any]:
        return {
            "id": self.element_id,
            "type": self.element_type,
            "tech_name": self.tech_name,
            "label": self.label,
            "value": self.value,
            "enabled": self.enabled,
            "visible": self.visible,
            "children": [child.to_dict() for child in self.children],
        }


@dataclass(slots=True)
class WindowSnapshot:
    """Structured representation of a window produced by ``dump_gui_structure``."""

    window_id: str
    title: str
    status_bar: str | None
    containers: Sequence[ElementState]
    elements: Sequence[ElementState]

    def to_dict(self) -> MutableMapping[str, Any]:
        return {
            "window_id": self.window_id,
            "title": self.title,
            "status_bar": self.status_bar,
            "containers": [container.to_dict() for container in self.containers],
            "elements": [element.to_dict() for element in self.elements],
        }


def flatten_elements(structure: Iterable[ElementState]) -> list[ElementState]:
    """Flatten element hierarchy into a list preserving traversal order."""

    result: list[ElementState] = []
    stack: list[ElementState] = list(structure)
    while stack:
        node = stack.pop()
        result.append(node)
        stack.extend(reversed(node.children))
    return list(reversed(result))
