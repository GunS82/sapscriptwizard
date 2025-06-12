"""Optional plugin providing semantic locator support."""

from .element_finder import SapElementFinder, DEFAULT_TARGET_TYPES
from .locator_helpers import (
    Position,
    ElementInfo,
    LocatorStrategy,
    ContentLocator,
    HLabelLocator,
    VLabelLocator,
    HLabelVLabelLocator,
    HLabelHLabelLocator,
)

__all__ = [
    "SapElementFinder",
    "DEFAULT_TARGET_TYPES",
    "Position",
    "ElementInfo",
    "LocatorStrategy",
    "ContentLocator",
    "HLabelLocator",
    "VLabelLocator",
    "HLabelVLabelLocator",
    "HLabelHLabelLocator",
]

