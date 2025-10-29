"""Deprecated window module."""
from __future__ import annotations

import warnings

from sapscriptwizard.compat.legacy import Window as _LegacyWindow

warnings.warn(
    "Importing window.Window is deprecated; use sapscriptwizard.core.window.Window",
    DeprecationWarning,
    stacklevel=2,
)

Window = _LegacyWindow

__all__ = ["Window"]
