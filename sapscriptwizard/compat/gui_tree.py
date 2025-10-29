"""Legacy GuiTree adapter."""
from __future__ import annotations

from ..gui.controls.tree import GuiTree as NewGuiTree
from . import warn_deprecated


class GuiTree(NewGuiTree):
    def __init__(self, session_handle, element: str) -> None:
        warn_deprecated()
        from ..core.session import SapSession

        legacy_session = SapSession(session_handle)
        window = legacy_session.window()
        super().__init__(window, element)


__all__ = ["GuiTree"]
