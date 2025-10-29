"""Deprecated GUI tree module."""
from __future__ import annotations

import warnings

from sapscriptwizard.gui.controls.tree import GuiTree as _GuiTree

warnings.warn(
    "Importing GuiTree from gui_tree.py is deprecated; use sapscriptwizard.gui.controls.tree",
    DeprecationWarning,
    stacklevel=2,
)

GuiTree = _GuiTree

__all__ = ["GuiTree"]
