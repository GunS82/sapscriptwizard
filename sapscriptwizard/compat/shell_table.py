"""Legacy ShellTable adapter."""
from __future__ import annotations

from ..gui.controls.table import ShellTable as NewShellTable
from . import warn_deprecated


class ShellTable(NewShellTable):
    def __init__(self, session_handle, element: str, load_table: bool = True) -> None:
        warn_deprecated()
        # ``session_handle`` exposes ``findById`` which :class:`Window` expects.
        from ..core.session import SapSession

        legacy_session = SapSession(session_handle)
        window = legacy_session.window()
        super().__init__(window, element)


__all__ = ["ShellTable"]
