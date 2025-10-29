"""Helper utilities to explain GUI elements."""
from __future__ import annotations

from ..core.types import ElementInfo, LocatorHint
from ..gui.finders.semantic_finder import SemanticFinder


def explain_id(structure: dict[str, object], element_id: str) -> ElementInfo:
    def _find(info: ElementInfo) -> ElementInfo | None:
        if info.element_id == element_id:
            return info
        for child in info.children:
            result = _find(child)
            if result:
                return result
        return None

    def _from_dict(data: dict[str, object]) -> ElementInfo:
        info = ElementInfo(
            element_id=data.get("id", ""),
            type=data.get("type", ""),
            tech_name=data.get("tech_name"),
            label=data.get("label"),
            value=data.get("value"),
            enabled=data.get("enabled"),
            visible=data.get("visible"),
        )
        info.children = [_from_dict(child) for child in data.get("children", [])]
        return info

    root = ElementInfo(
        element_id=structure.get("window_id", ""),
        type="GuiWindow",
        label=structure.get("title", ""),
        children=[_from_dict(child) for child in structure.get("containers", [])],
    )
    match = _find(root)
    if not match:
        raise ValueError(f"Element {element_id} not found in structure")
    return match


def suggest_locator(element: ElementInfo) -> list[LocatorHint]:
    finder = SemanticFinder(element)
    return finder.suggest_locator(element)
