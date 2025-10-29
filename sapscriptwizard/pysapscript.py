"""Deprecated entry point. Use :mod:`sapscriptwizard.cli.sapwiz`."""
from __future__ import annotations

from .compat import warn_deprecated
from .cli.sapwiz import main

warn_deprecated()

if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
