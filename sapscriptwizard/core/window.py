"""Window level automation primitives."""
from __future__ import annotations

import logging
from dataclasses import dataclass
from typing import Any, Iterable, Iterator, Sequence

from .com_gateway import ComGateway
from .errors import SapError, SapErrorCode
from .types import ElementState, WindowSnapshot


_LOGGER = logging.getLogger(__name__)


def _iter_children(container: Any) -> Iterator[Any]:
    children = getattr(container, "Children", None)
    if children is None:
        return iter(())
    if isinstance(children, Sequence):
        return iter(children)
    if hasattr(children, "Count") and hasattr(children, "Item"):
        return (children.Item(idx) for idx in range(children.Count))
    if hasattr(children, "__iter__"):
        return iter(children)
    return iter(())


@dataclass(slots=True)
class ElementProxy:
    element_id: str
    raw: Any

    @property
    def text(self) -> str:
        return getattr(self.raw, "Text", getattr(self.raw, "text", ""))


class Window:
    """Facade for interacting with a single SAP GUI window."""

    def __init__(self, session: "SapSession", window_id: str) -> None:
        from .session import SapSession

        if not isinstance(session, SapSession):  # pragma: no cover - defensive
            raise TypeError("session must be SapSession")
        self._session = session
        self._window_id = window_id
        self._logger = logging.getLogger(f"{__name__}.{window_id}")

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    @property
    def gateway(self) -> ComGateway:
        return self._session.gateway

    def _qualify_id(self, element_id: str) -> str:
        if element_id.startswith("wnd["):
            return element_id
        if element_id.startswith("/"):
            return f"{self._window_id}{element_id}"
        return f"{self._window_id}/{element_id}"

    def _get_window(self) -> Any:
        return self.gateway.call(
            getattr(self._session.raw, "findById"),
            self._window_id,
            error_code=SapErrorCode.NOT_FOUND,
        )

    def _get_element(self, element_id: str) -> ElementProxy:
        qualified_id = self._qualify_id(element_id)
        try:
            raw = self.gateway.call(
                getattr(self._session.raw, "findById"),
                qualified_id,
                error_code=SapErrorCode.NOT_FOUND,
            )
        except SapError as exc:
            self._logger.error("Element %s not found", qualified_id)
            raise
        return ElementProxy(qualified_id, raw)

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------
    def find_by_id(self, element_id: str) -> Any:
        return self._get_element(element_id).raw

    def read_text(self, element_id: str) -> str:
        proxy = self._get_element(element_id)
        value = getattr(proxy.raw, "Value", getattr(proxy.raw, "value", None))
        if value is None:
            value = getattr(proxy.raw, "Text", getattr(proxy.raw, "text", ""))
        return str(value)

    def write(self, element_id: str, text: str) -> None:
        proxy = self._get_element(element_id)
        enabled = getattr(proxy.raw, "Enabled", True)
        if not enabled:
            raise SapError(
                SapErrorCode.DISABLED,
                f"Element {proxy.element_id} is disabled",
            )
        if hasattr(proxy.raw, "Value"):
            setattr(proxy.raw, "Value", text)
        elif hasattr(proxy.raw, "value"):
            setattr(proxy.raw, "value", text)
        elif hasattr(proxy.raw, "Text"):
            setattr(proxy.raw, "Text", text)
        elif hasattr(proxy.raw, "text"):
            setattr(proxy.raw, "text", text)
        else:
            raise SapError(SapErrorCode.UNSUPPORTED, f"Element {proxy.element_id} is not writable")

    def press(self, element_id: str) -> None:
        proxy = self._get_element(element_id)
        handler = getattr(proxy.raw, "press", None) or getattr(proxy.raw, "Press", None)
        if not callable(handler):
            raise SapError(SapErrorCode.UNSUPPORTED, f"Element {proxy.element_id} cannot be pressed")
        self.gateway.call(handler, error_code=SapErrorCode.COM_FAIL)

    def send_vkey(self, code: int) -> None:
        window = self._get_window()
        handler = getattr(window, "sendVKey", None)
        if not callable(handler):
            raise SapError(SapErrorCode.UNSUPPORTED, "Window does not support sendVKey")
        self.gateway.call(handler, code, error_code=SapErrorCode.COM_FAIL)

    # ------------------------------------------------------------------
    # Inspection
    # ------------------------------------------------------------------
    def dump_gui_structure(self, max_depth: int = 2, max_children: int = 100) -> dict[str, Any]:
        window = self._get_window()
        snapshot = self._build_snapshot(window, max_depth=max_depth, max_children=max_children)
        return snapshot.to_dict()

    def _build_snapshot(
        self,
        window: Any,
        *,
        max_depth: int,
        max_children: int,
    ) -> WindowSnapshot:
        title = getattr(window, "Text", getattr(window, "text", ""))
        status_bar = getattr(window, "StatusBar", None)
        status_text = getattr(status_bar, "Text", None) if status_bar else None
        containers: list[ElementState] = []
        elements: list[ElementState] = []

        for child in self._collect_children(window, depth=0, max_depth=max_depth, max_children=max_children):
            (containers if child.element_type.lower().endswith("container") else elements).append(child)

        return WindowSnapshot(
            window_id=self._window_id,
            title=str(title),
            status_bar=str(status_text) if status_text else None,
            containers=containers,
            elements=elements,
        )

    def _collect_children(
        self,
        parent: Any,
        *,
        depth: int,
        max_depth: int,
        max_children: int,
    ) -> Iterable[ElementState]:
        if depth >= max_depth:
            return []
        collected: list[ElementState] = []
        for idx, raw_child in enumerate(_iter_children(parent)):
            if idx >= max_children:
                break
            element_id = getattr(raw_child, "Id", getattr(raw_child, "ID", None))
            element_type = getattr(raw_child, "Type", getattr(raw_child, "type", ""))
            tech_name = getattr(raw_child, "TechnicalName", getattr(raw_child, "Name", None))
            label = getattr(raw_child, "Text", getattr(raw_child, "Tooltip", None))
            value = getattr(raw_child, "Value", getattr(raw_child, "Text", None))
            enabled = bool(getattr(raw_child, "Enabled", True))
            visible = bool(getattr(raw_child, "Visible", True))
            children = list(
                self._collect_children(
                    raw_child,
                    depth=depth + 1,
                    max_depth=max_depth,
                    max_children=max_children,
                )
            )
            collected.append(
                ElementState(
                    element_id=element_id or "",
                    element_type=str(element_type),
                    tech_name=str(tech_name) if tech_name else None,
                    label=str(label) if label else None,
                    value=value,
                    enabled=enabled,
                    visible=visible,
                    children=children,
                )
            )
        return collected

    # Finders --------------------------------------------------------------
    @property
    def finder(self) -> "SemanticFinder":
        from ..gui.finders.semantic_finder import SemanticFinder

        return SemanticFinder(self)


__all__ = ["Window", "ElementProxy"]
