"""Recipes for working with transaction SE80."""
from __future__ import annotations

from ..core.window import Window
from ..gui.controls.abap_editor import GuiAbapEditor


class SE80:
    def __init__(self, window: Window) -> None:
        self.window = window

    def read_source(self, editor_id: str) -> str:
        editor = GuiAbapEditor(self.window.find_by_id(editor_id))
        return editor.get_all_text()
