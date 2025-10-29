"""Legacy Window adapter."""
from __future__ import annotations

from ..core.session import SapSession
from ..core.window import Window as NewWindow
from . import warn_deprecated


class LegacyWindow(NewWindow):
    def __init__(
        self,
        application,
        connection,
        connection_handle,
        session,
        session_handle,
    ) -> None:
        warn_deprecated()
        self.application = application
        self.connection = connection
        self.connection_handle = connection_handle
        self.session = session
        self.session_handle = session_handle
        self._legacy_session = SapSession(session_handle)
        super().__init__(self._legacy_session, f"wnd[{session}]")


Window = LegacyWindow


__all__ = ["LegacyWindow", "Window"]
