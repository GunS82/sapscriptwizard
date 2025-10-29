"""Deprecated element finder module."""
from __future__ import annotations

import warnings

from sapscriptwizard.gui.finders.semantic_finder import SemanticFinder

warnings.warn(
    "element_finder is deprecated; use sapscriptwizard.gui.finders",
    DeprecationWarning,
    stacklevel=2,
)

__all__ = ["SemanticFinder"]
