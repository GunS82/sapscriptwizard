"""Tree control helpers."""
from __future__ import annotations

from typing import Any, Iterable, Iterator, Sequence

from ...core.errors import SapError, SapErrorCode


class GuiTree:
    def __init__(self, window: "Window", element_id: str) -> None:
        from ...core.window import Window

        if not isinstance(window, Window):  # pragma: no cover - defensive
            raise TypeError("window must be Window")
        self._window = window
        self._element_id = element_id
        self._raw = window.find_by_id(element_id)

    @property
    def raw(self) -> Any:
        return self._raw

    def expand_path(self, path: Sequence[str]) -> None:
        node_id = self._ensure_node(path)
        expand = getattr(self._raw, "ExpandNode", None)
        if not callable(expand):
            raise SapError(SapErrorCode.UNSUPPORTED, "Tree does not support ExpandNode")
        self._window.gateway.call(expand, node_id)

    def select_path(self, path: Sequence[str]) -> None:
        node_id = self._ensure_node(path)
        select = getattr(self._raw, "SelectNode", None)
        if not callable(select):
            raise SapError(SapErrorCode.UNSUPPORTED, "Tree does not support SelectNode")
        self._window.gateway.call(select, node_id)

    def _ensure_node(self, path: Sequence[str]) -> str:
        path_text = "\x7c".join(path)
        getter = getattr(self._raw, "GetNodeKey", None)
        if not callable(getter):
            raise SapError(SapErrorCode.UNSUPPORTED, "Tree does not support GetNodeKey")
        key = getter(path_text)
        if key is None:
            raise SapError(SapErrorCode.NOT_FOUND, f"Tree path {'/'.join(path)} not found")
        return key

    def iter_nodes(self) -> Iterator[str]:
        collection = getattr(self._raw, "Nodes", None)
        if collection is None:
            return iter(())
        if hasattr(collection, "Count") and hasattr(collection, "Item"):
            return (collection.Item(idx) for idx in range(collection.Count))
        if isinstance(collection, Sequence):
            return iter(collection)
        if hasattr(collection, "__iter__"):
            return iter(collection)
        return iter(())


__all__ = ["GuiTree"]
