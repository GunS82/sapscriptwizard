"""WE19 transaction helpers."""
from __future__ import annotations

from pathlib import Path

from ..core.window import Window
from ..gui.controls.text_editor import GuiTextEditor


class WE19:
    def __init__(self, window: Window) -> None:
        self.window = window

    def upload_xml(self, editor_id: str, xml_path: str | Path) -> WE19:
        editor = GuiTextEditor(self.window.find_by_id(editor_id))
        content = Path(xml_path).read_text(encoding="utf-8")
        editor.set_all_text(content)
        return self

    def execute(self, button_id: str) -> None:
        self.window.press(button_id)
