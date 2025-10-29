"""Deprecated locator helpers."""
from __future__ import annotations

import warnings

from sapscriptwizard.gui.finders.semantic_finder import SemanticFinder

warnings.warn(
    "locator_helpers is deprecated; use sapscriptwizard.gui.finders.semantic_finder",
    DeprecationWarning,
    stacklevel=2,
)

__all__ = ["SemanticFinder"]
