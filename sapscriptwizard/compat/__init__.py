"""Compatibility layer for the legacy 0.x API."""
from __future__ import annotations

import warnings

_DEPRECATION_MESSAGE = (
    "The legacy API is deprecated and will be removed in the next major release. "
    "Consult MIGRATION.md for guidance."
)


def warn_deprecated() -> None:
    warnings.warn(_DEPRECATION_MESSAGE, DeprecationWarning, stacklevel=2)


__all__ = ["warn_deprecated"]
