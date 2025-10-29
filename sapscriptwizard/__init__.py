"""sapscriptwizard package public API."""
from __future__ import annotations

import json
import logging

from .compat.legacy import Sapscript as LegacySapscript
from .core.session import SapSession
from .core.window import Window
from .gui.controls.abap_editor import GuiAbapEditor
from .gui.controls.table import ShellTable
from .gui.controls.text_editor import GuiTextEditor
from .gui.controls.tree import GuiTree
from .helpers.explain import explain_id, suggest_locator

__all__ = [
    "SapSession",
    "Window",
    "ShellTable",
    "GuiTree",
    "GuiAbapEditor",
    "GuiTextEditor",
    "LegacySapscript",
    "explain_id",
    "suggest_locator",
]


def configure_logging(level: int = logging.INFO) -> None:
    """Configure package-wide logging to emit JSON formatted records."""

    class JsonFormatter(logging.Formatter):
        def format(self, record: logging.LogRecord) -> str:  # pragma: no cover - formatting only
            payload = {
                "level": record.levelname,
                "name": record.name,
                "message": record.getMessage(),
            }
            if record.exc_info:
                payload["exception"] = self.formatException(record.exc_info)
            return json.dumps(payload, ensure_ascii=False)

    handler = logging.StreamHandler()
    handler.setFormatter(JsonFormatter())
    root = logging.getLogger("sapscriptwizard")
    root.setLevel(level)
    root.handlers.clear()
    root.addHandler(handler)
