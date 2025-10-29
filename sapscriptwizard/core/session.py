"""Session management for SAP GUI scripting."""
from __future__ import annotations

import contextlib
import logging
from dataclasses import dataclass
from typing import Any, Callable

from .com_gateway import ComGateway, RetryPolicy
from .errors import SapError, SapErrorCode


_LOGGER = logging.getLogger(__name__)


def _default_retry_exceptions() -> tuple[type[Exception], ...]:
    try:  # pragma: no cover - platform specific
        import pywintypes  # type: ignore

        return (pywintypes.com_error,)
    except Exception:  # noqa: BLE001 - fall back to generic runtime
        return ()


@dataclass(slots=True)
class SessionInfo:
    connection_id: int
    session_id: int
    system_name: str | None = None
    client: str | None = None


class SapSession:
    """High level session wrapper for SAP GUI scripting."""

    def __init__(
        self,
        raw_session: Any,
        *,
        com_gateway: ComGateway | None = None,
        default_timeout: float = 30.0,
        logger: logging.Logger | None = None,
        info: SessionInfo | None = None,
    ) -> None:
        self._raw_session = raw_session
        self._gateway = com_gateway or ComGateway(
            retry_policy=RetryPolicy(max_attempts=5, base_delay=0.1, multiplier=1.8),
            retry_exceptions=_default_retry_exceptions(),
        )
        self._default_timeout = default_timeout
        self._logger = logger or _LOGGER
        self.info = info

    # ------------------------------------------------------------------
    # Construction helpers
    # ------------------------------------------------------------------
    @classmethod
    def from_gui(  # pragma: no cover - integration helper
        cls,
        connection_index: int = 0,
        session_index: int = 0,
        *,
        com_gateway: ComGateway | None = None,
        default_timeout: float = 30.0,
        logger: logging.Logger | None = None,
    ) -> "SapSession":
        """Create a :class:`SapSession` from a running SAP GUI instance."""

        try:
            import win32com.client  # type: ignore[import-not-found]
        except ImportError as exc:  # pragma: no cover - executed on Windows only
            raise SapError(SapErrorCode.UNSUPPORTED, "win32com is required for live sessions") from exc

        gui_auto = win32com.client.GetObject("SAPGUI")
        if gui_auto is None:
            raise SapError(SapErrorCode.NOT_FOUND, "SAP GUI automation instance not found")

        application = gui_auto.GetScriptingEngine
        connection = application.Children(connection_index)
        session = connection.Children(session_index)

        info = SessionInfo(
            connection_id=connection_index,
            session_id=session_index,
            system_name=getattr(connection, "SystemName", None),
            client=getattr(connection, "Client", None),
        )
        return cls(
            session,
            com_gateway=com_gateway,
            default_timeout=default_timeout,
            logger=logger,
            info=info,
        )

    # ------------------------------------------------------------------
    # Context manager support
    # ------------------------------------------------------------------
    def __enter__(self) -> "SapSession":  # pragma: no cover - trivial
        return self

    def __exit__(self, exc_type, exc, tb) -> None:  # pragma: no cover - trivial
        self.close()

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------
    @property
    def gateway(self) -> ComGateway:
        return self._gateway

    def close(self) -> None:
        """Close the underlying session if possible."""

        if self._raw_session is None:
            return
        with contextlib.suppress(Exception):
            close = getattr(self._raw_session, "CloseConnection", None)
            if callable(close):
                close()
        self._raw_session = None

    def ensure_alive(self) -> None:
        """Ensure the session is still connected, raising :class:`SapError` otherwise."""

        if self._raw_session is None:
            raise SapError(SapErrorCode.TIMEOUT, "Session already closed")

    def window(self, window_id: str = "wnd[0]") -> "Window":
        """Return a :class:`Window` proxy for ``window_id``."""

        from .window import Window

        self.ensure_alive()
        return Window(self, window_id)

    # Convenience wrappers -------------------------------------------------
    def press(self, element_id: str) -> None:
        self.window().press(element_id)

    def read_text(self, element_id: str) -> str:
        return self.window().read_text(element_id)

    def write(self, element_id: str, text: str) -> None:
        self.window().write(element_id, text)

    def run(self, func: Callable[[Any], Any]) -> Any:
        """Execute ``func`` with the raw session via :class:`ComGateway`."""

        self.ensure_alive()
        return self._gateway.call(func, self._raw_session)

    @property
    def raw(self) -> Any:
        self.ensure_alive()
        return self._raw_session
