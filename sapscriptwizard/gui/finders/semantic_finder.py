"""Semantic locator resolution."""
from __future__ import annotations

import re
from collections.abc import Iterable

from ...core.errors import ErrorCode, SapError
from ...core.types import ElementInfo, LocatorHint, LocatorHintType
from .locator_model import LocatorQuery

_LABEL_PATTERN = re.compile(r'@"(?P<label>[^"]+)"')
_TITLE_PATTERN = re.compile(r'="(?P<title>[^"]+)"')
_ID_PATTERN = re.compile(r'id:"(?P<id>[^"]+)"')
_COL_PATTERN = re.compile(r'col:"(?P<column>[^"]+)"\s+row:(?P<row>\d+)')


class SemanticFinder:
    """Resolve human readable queries into element identifiers."""

    def __init__(self, root: ElementInfo) -> None:
        self.root = root

    def find(self, query: str) -> str:
        locator = LocatorQuery(query)
        if match := _ID_PATTERN.search(locator.raw):
            return match.group("id")

        if match := _TITLE_PATTERN.search(locator.raw):
            title = match.group("title")
            element = self._find_by_attribute("label", title)
            if element:
                return element.element_id

        if match := _LABEL_PATTERN.search(locator.raw):
            label = match.group("label")
            element = self._find_by_attribute("label", label)
            if element:
                return element.element_id

        if match := _COL_PATTERN.search(locator.raw):
            return f"col:{match.group('column')} row:{match.group('row')}"

        raise SapError(ErrorCode.NOT_FOUND, f"Could not resolve locator: {query}")

    def _find_by_attribute(self, attr: str, value: str) -> ElementInfo | None:
        for element in self._iter_elements(self.root):
            if getattr(element, attr) == value:
                return element
        return None

    def _iter_elements(self, root: ElementInfo) -> Iterable[ElementInfo]:
        yield root
        for child in root.children:
            yield from self._iter_elements(child)

    def suggest_locator(self, element: ElementInfo) -> list[LocatorHint]:
        hints: list[LocatorHint] = []
        if element.label:
            hints.append(LocatorHint(LocatorHintType.BY_TITLE, {"title": element.label}))
        if element.value:
            hints.append(LocatorHint(LocatorHintType.H_LABEL, {"label": element.value}))
        return hints
