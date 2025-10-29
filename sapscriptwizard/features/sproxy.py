"""SPROXY transaction helpers."""
from __future__ import annotations

from ..core.window import Window
from ..gui.controls.tree import GuiTree


class SPROXY:
    def __init__(self, window: Window) -> None:
        self.window = window

    def navigate_repository(self, tree_id: str, path: list[str]) -> None:
        tree = GuiTree(self.window.find_by_id(tree_id))
        tree.select_path(path)
