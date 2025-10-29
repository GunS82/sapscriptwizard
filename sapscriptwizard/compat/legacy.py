"""Compatibility layer exposing the legacy API."""
from __future__ import annotations

import warnings
from typing import Any

from ..core.session import SapSession
from ..core.window import Window as NewWindow
from ..gui.controls.table import ShellTable as NewShellTable

_DEPRECATION = "Legacy API will be removed in the next major release; see MIGRATION.md"


class Sapscript:
    def __init__(self) -> None:
        warnings.warn(_DEPRECATION, DeprecationWarning, stacklevel=2)
        self._session = SapSession()

    def launch_sap(self, *args: Any, **kwargs: Any) -> None:  # pragma: no cover - passthrough
        warnings.warn(_DEPRECATION, DeprecationWarning, stacklevel=2)
        self._session.connect()

    def attach(self) -> None:
        warnings.warn(_DEPRECATION, DeprecationWarning, stacklevel=2)
        self._session.connect()

    def get_session(self) -> SapSession:
        return self._session

    def get_window(self, index: int = 0) -> Window:
        return Window(self._session.window(index))


class Window:
    def __init__(self, window: NewWindow) -> None:
        warnings.warn(_DEPRECATION, DeprecationWarning, stacklevel=2)
        self._window = window

    def findById(self, element_id: str) -> Any:  # noqa: N802 - legacy casing
        return self._window.find_by_id(element_id)

    def read_text(self, element_id: str) -> str:
        return self._window.read_text(element_id)

    def write(self, element_id: str, text: str) -> None:
        self._window.write(element_id, text)

    def press(self, element_id: str) -> None:
        self._window.press(element_id)

    def send_vkey(self, code: int) -> None:
        self._window.send_vkey(code)

    def dump_gui_structure(self, max_depth: int = 2, max_children: int = 100) -> dict[str, Any]:
        return self._window.dump_gui_structure(max_depth=max_depth, max_children=max_children)


class ShellTable(NewShellTable):
    def __init__(self, com_table: Any) -> None:
        warnings.warn(_DEPRECATION, DeprecationWarning, stacklevel=2)
        super().__init__(com_table)


__all__ = ["Sapscript", "Window", "ShellTable"]
