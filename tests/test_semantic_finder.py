from __future__ import annotations

import pytest

from sapscriptwizard.core.errors import SapError
from sapscriptwizard.core.types import ElementInfo
from sapscriptwizard.gui.finders.semantic_finder import SemanticFinder


def build_tree() -> ElementInfo:
    child = ElementInfo(element_id="child", type="GuiButton", label="Execute")
    root = ElementInfo(element_id="root", type="GuiWindow", label="Main", children=[child])
    return root


def test_find_by_title():
    finder = SemanticFinder(build_tree())
    assert finder.find('="Execute"') == "child"


def test_find_not_found():
    finder = SemanticFinder(build_tree())
    with pytest.raises(SapError):
        finder.find('="Missing"')
