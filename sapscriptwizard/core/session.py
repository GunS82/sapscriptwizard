"""High level session management for SAP GUI scripting."""
from __future__ import annotations

import atexit
import logging
from dataclasses import dataclass
from typing import Any

try:  # pragma: no cover - optional dependency for non-Windows tests
    import win32com.client  # type: ignore
except Exception:  # pragma: no cover
    win32com = None  # type: ignore

from .com_gateway import ComGateway, RetryPolicy
from .errors import ErrorCode, SapError
from .window import Window

log = logging.getLogger(__name__)


@dataclass(slots=True)
class SessionConfig:
    connection_index: int = 0
    session_index: int = 0
    auto_close: bool = True


class SapSession:
    """Entry point to SAP GUI scripting."""

    def __init__(
        self,
        *,
        gateway: ComGateway | None = None,
        config: SessionConfig | None = None,
    ) -> None:
        self.gateway = gateway or ComGateway(RetryPolicy())
        self.config = config or SessionConfig()
        self._application: Any | None = None
        self._session: Any | None = None

    def connect(self) -> None:
        """Attach to a running SAP GUI session."""

        if win32com is None:  # pragma: no cover - platform safeguard
            raise SapError(
                ErrorCode.COM_FAIL,
                "win32com is unavailable; SAP GUI scripting requires Windows",
            )

        try:
            log.info("Connecting to SAP GUI scripting engine")
            sapgui = win32com.client.GetObject("SAPGUI")  # type: ignore[attr-defined]
            self._application = sapgui.GetScriptingEngine
        except Exception as exc:  # pragma: no cover - external dependency
            raise SapError(
                ErrorCode.COM_FAIL,
                "Could not access SAP GUI scripting engine",
                str(exc),
            ) from exc

        connections = self._application.Connections
        connection = self.gateway.call_method(connections, "Item", self.config.connection_index)
        self._session = connection.Children(self.config.session_index)

        if self.config.auto_close:
            atexit.register(self.close)

    @property
    def raw_session(self) -> Any:
        if self._session is None:
            raise SapError(ErrorCode.COM_FAIL, "Session is not connected")
        return self._session

    def window(self, index: int = 0) -> Window:
        """Return the :class:`Window` at *index*."""

        win = self.gateway.call_method(self.raw_session.Children, "Item", index)
        return Window(win, gateway=self.gateway)

    def close(self) -> None:
        """Attempt to close the session."""

        if self._session is None:
            return
        try:
            self.gateway.call_method(self._session, "Close")
        except SapError:
            log.warning("Failed to close SAP session cleanly", exc_info=True)
        finally:
            self._session = None


__all__ = ["SapSession", "SessionConfig"]
