"""SProxy (Enterprise Service Browser) helpers."""
from __future__ import annotations

from ..core.session import SapSession
from ..gui.controls.tree import GuiTree


class SProxy:
    def __init__(self, session: SapSession) -> None:
        self._session = session

    def open_service(self, service_name: str) -> None:
        window = self._session.window()
        tree = GuiTree(window, "usr/treeServices")
        tree.select_path([service_name])
        window.press("usr/btnDisplay")


__all__ = ["SProxy"]
