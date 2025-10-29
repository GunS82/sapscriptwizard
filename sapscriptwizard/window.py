"""Compatibility shim for the legacy :mod:`window` module."""
from __future__ import annotations

from .compat.window import LegacyWindow, Window

__all__ = ["Window", "LegacyWindow"]
