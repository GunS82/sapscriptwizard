"""Abstractions for GuiTree controls."""
from __future__ import annotations

from typing import Any

from ...core.com_gateway import ComGateway
from ...core.errors import ErrorCode, SapError


class GuiTree:
    """Wrapper around GuiTree control."""

    def __init__(self, com_tree: Any, *, gateway: ComGateway | None = None) -> None:
        self._tree = com_tree
        self.gateway = gateway or ComGateway()

    def get_selected_node(self) -> str:
        try:
            return str(self.gateway.get_attr(self._tree, "SelectedNode"))
        except SapError:
            raise
        except Exception as exc:
            raise SapError(
                ErrorCode.UNSUPPORTED,
                "Tree does not expose SelectedNode",
                str(exc),
            ) from exc

    def select_path(self, path: list[str]) -> None:
        try:
            for segment in path:
                self.gateway.call_method(self._tree, "ExpandNode", segment)
            self.gateway.call_method(self._tree, "SelectNode", path[-1])
        except SapError:
            raise
        except Exception as exc:
            raise SapError(ErrorCode.NOT_FOUND, "Tree path selection failed", str(exc)) from exc

    def to_dict(self) -> dict[str, Any]:
        return {
            "selected": self.get_selected_node(),
            "visible_nodes": list(self.iter_visible_nodes()),
        }

    def iter_visible_nodes(self) -> list[str]:
        try:
            node_keys = self.gateway.get_attr(self._tree, "GetAllNodeKeys")
        except SapError:
            raise
        except Exception:
            node_keys = []

        if callable(node_keys):  # pragma: no cover - optional path
            node_keys = node_keys()

        return [str(key) for key in node_keys]
