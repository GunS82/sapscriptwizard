"""High level helpers for WE19 IDoc testing."""
from __future__ import annotations

from pathlib import Path

from ..core.session import SapSession


class WE19:
    def __init__(self, session: SapSession) -> None:
        self._session = session

    def upload_xml(self, path: str | Path) -> "WE19":
        window = self._session.window()
        window.press("usr/btnImport")
        window.write("usr/txtFilePath", str(Path(path)))
        window.press("usr/btnExecute")
        return self

    def execute(self) -> None:
        self._session.window().press("usr/btnRun")


__all__ = ["WE19"]
