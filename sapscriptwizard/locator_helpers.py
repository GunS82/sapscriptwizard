"""Legacy helpers kept for import compatibility."""
from __future__ import annotations

from .compat import warn_deprecated
from .helpers.explain import explain_id, suggest_locator

warn_deprecated()

__all__ = ["explain_id", "suggest_locator"]
