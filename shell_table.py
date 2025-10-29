"""Deprecated shell table module."""
from __future__ import annotations

import warnings

from sapscriptwizard.compat.legacy import ShellTable as _LegacyShellTable

warnings.warn(
    (
        "Importing ShellTable from the top level is deprecated; "
        "use sapscriptwizard.gui.controls.table.ShellTable"
    ),
    DeprecationWarning,
    stacklevel=2,
)

ShellTable = _LegacyShellTable

__all__ = ["ShellTable"]
