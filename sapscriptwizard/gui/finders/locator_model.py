"""Locator model definitions used by semantic finder."""
from __future__ import annotations

from dataclasses import dataclass


@dataclass(slots=True)
class LocatorQuery:
    raw: str

    def normalised(self) -> str:
        return self.raw.strip()
