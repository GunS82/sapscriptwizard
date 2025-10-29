"""Compatibility shim mapping to the new semantic finder."""
from __future__ import annotations

from .compat import warn_deprecated
from .gui.finders.semantic_finder import SemanticFinder


class SapElementFinder(SemanticFinder):
    def __init__(self, session_handle) -> None:
        warn_deprecated()
        from .core.session import SapSession

        legacy_session = SapSession(session_handle)
        window = legacy_session.window()
        super().__init__(window)


__all__ = ["SapElementFinder", "SemanticFinder"]
