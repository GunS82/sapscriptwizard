"""Deprecated type aliases."""
from __future__ import annotations

import warnings

warnings.warn(
    "types_ package is deprecated; use sapscriptwizard.core.errors and sapscriptwizard.core.types",
    DeprecationWarning,
    stacklevel=2,
)

from .exceptions import *  # noqa: F401,F403
from .types import *  # noqa: F401,F403

__all__ = []  # populated by star imports
