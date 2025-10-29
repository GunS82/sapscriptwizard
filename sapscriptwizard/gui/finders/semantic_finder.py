"""Semantic locator resolution for SAP GUI windows."""
from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, Iterable, Iterator

from .locator_model import ElementInfo, LocatorHint, LocatorType


_ID_PATTERN = re.compile(r'id:"(?P<id>[^"]+)"')
_TITLE_PATTERN = re.compile(r'^="(?P<title>.+)"$')
_LABEL_PATTERN = re.compile(r'^@"(?P<label>.+)"$')
_CELL_PATTERN = re.compile(r'col:"(?P<col>[^"]+)"\s+row:(?P<row>\d+)')


@dataclass(slots=True)
class _FlattenedElement:
    info: ElementInfo
    raw: dict[str, Any]


class SemanticFinder:
    def __init__(self, window: "Window", *, max_depth: int = 3, max_children: int = 200) -> None:
        self._window = window
        self._max_depth = max_depth
        self._max_children = max_children

    # ------------------------------------------------------------------
    def find(self, query: str) -> str:
        query = query.strip()
        if not query:
            raise ValueError("query is empty")

        if match := _ID_PATTERN.fullmatch(query):
            return match.group("id")
        if match := _TITLE_PATTERN.fullmatch(query):
            return self._find_by_title(match.group("title"))
        if match := _LABEL_PATTERN.fullmatch(query):
            return self._find_by_label(match.group("label"))
        if match := _CELL_PATTERN.fullmatch(query):
            return self._find_by_table_cell(match.group("col"), int(match.group("row")))

        raise ValueError(f"Unsupported query grammar: {query}")

    # ------------------------------------------------------------------
    def _fetch_elements(self) -> list[_FlattenedElement]:
        snapshot = self._window.dump_gui_structure(max_depth=self._max_depth, max_children=self._max_children)
        containers = snapshot.get("containers", [])
        elements = snapshot.get("elements", [])
        flattened: list[_FlattenedElement] = []
        for entry in self._iter_entries(containers + elements):
            info = ElementInfo(
                element_id=entry.get("id", ""),
                label=entry.get("label"),
                value=entry.get("value"),
                element_type=entry.get("type", ""),
                tech_name=entry.get("tech_name"),
            )
            flattened.append(_FlattenedElement(info, entry))
        return flattened

    def _iter_entries(self, entries: Iterable[dict[str, Any]]) -> Iterator[dict[str, Any]]:
        for entry in entries:
            yield entry
            for child in entry.get("children", []) or []:
                yield from self._iter_entries([child])

    def _find_by_label(self, label: str) -> str:
        candidates = self._fetch_elements()
        label_cf = label.casefold()
        matches = [c for c in candidates if (c.info.label or "").casefold() == label_cf]
        if not matches:
            raise LookupError(f"No element with label {label}")
        return matches[0].info.element_id

    def _find_by_title(self, title: str) -> str:
        candidates = self._fetch_elements()
        title_cf = title.casefold()
        for candidate in candidates:
            if (candidate.info.label or "").casefold() == title_cf:
                return candidate.info.element_id
            if (candidate.info.value or "").casefold() == title_cf:
                return candidate.info.element_id
        raise LookupError(f"No element with title {title}")

    def _find_by_table_cell(self, column: str, row: int) -> str:
        candidates = self._fetch_elements()
        column_cf = column.casefold()
        for candidate in candidates:
            value = candidate.raw.get("value")
            if isinstance(value, list):
                try:
                    row_data = value[row]
                except IndexError as exc:
                    raise LookupError(f"Row {row} is out of range") from exc
                if isinstance(row_data, dict):
                    for key, cell_value in row_data.items():
                        if key.casefold() == column_cf:
                            return f"{candidate.info.element_id}/row[{row}]/col[{key}]"
        raise LookupError(f"No table element with column {column}")

    # ------------------------------------------------------------------
    def suggest(self, info: ElementInfo) -> list[LocatorHint]:
        hints: list[LocatorHint] = []
        if info.label:
            hints.append(LocatorHint(LocatorType.H_LABEL, f'@"{info.label}"', confidence=0.9))
        if info.value:
            hints.append(LocatorHint(LocatorType.BY_TITLE, f'="{info.value}"', confidence=0.7))
        hints.append(LocatorHint(LocatorType.BY_ID, f'id:"{info.element_id}"', confidence=0.5))
        return hints


__all__ = ["SemanticFinder"]
