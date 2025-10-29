"""Helper functions that explain GUI element structure and suggest locators."""
from __future__ import annotations

from typing import Any, Iterable, Iterator

from ..core.window import Window
from ..gui.finders.locator_model import ElementInfo, LocatorHint


def explain_id(window: Window, element_id: str) -> ElementInfo:
    structure = window.dump_gui_structure(max_depth=5, max_children=500)
    for entry in _iter_entries(structure.get("containers", []) + structure.get("elements", [])):
        if entry.get("id") == element_id:
            return ElementInfo(
                element_id=element_id,
                label=entry.get("label"),
                value=entry.get("value"),
                element_type=entry.get("type", ""),
                tech_name=entry.get("tech_name"),
            )
    raise LookupError(f"Element {element_id} not found in window snapshot")


def suggest_locator(window: Window, element_id: str) -> list[LocatorHint]:
    info = explain_id(window, element_id)
    finder = window.finder
    return finder.suggest(info)


def _iter_entries(entries: Iterable[dict[str, Any]]) -> Iterator[dict[str, Any]]:
    for entry in entries:
        yield entry
        for child in entry.get("children", []) or []:
            yield from _iter_entries([child])


__all__ = ["explain_id", "suggest_locator"]
