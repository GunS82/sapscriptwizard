"""High level recipes for SE80 transactions."""
from __future__ import annotations

from ..core.session import SapSession
from ..gui.controls.abap_editor import GuiAbapEditor


class SE80:
    def __init__(self, session: SapSession) -> None:
        self._session = session

    def read_source(self, program_name: str) -> str:
        window = self._session.window()
        window.write("usr/txtProgram", program_name)
        window.press("usr/btnDisplay")
        editor = GuiAbapEditor(window, "usr/subEDITOR")
        return editor.get_all_text()


__all__ = ["SE80"]
