from __future__ import annotations

from sapscriptwizard.gui.finders.locator_model import ElementInfo, LocatorType
from sapscriptwizard.gui.finders.semantic_finder import SemanticFinder


class FakeWindow:
    def dump_gui_structure(self, max_depth: int = 2, max_children: int = 100):
        return {
            "containers": [],
            "elements": [
                {
                    "id": "wnd[0]/usr/txtUser",
                    "type": "GuiTextField",
                    "label": "User",
                    "value": "",
                    "children": [],
                },
                {
                    "id": "wnd[0]/usr/btnEXEC",
                    "type": "GuiButton",
                    "label": "Execute",
                    "value": "Execute",
                    "children": [],
                },
                {
                    "id": "wnd[0]/usr/cntlGRID",
                    "type": "GuiShell",
                    "label": "Table",
                    "value": [
                        {"Status": "OK"},
                        {"Status": "WARN"},
                    ],
                    "children": [],
                },
            ],
        }


def test_find_by_label():
    finder = SemanticFinder(FakeWindow())
    assert finder.find('@"User"') == "wnd[0]/usr/txtUser"


def test_find_by_title():
    finder = SemanticFinder(FakeWindow())
    assert finder.find('="Execute"') == "wnd[0]/usr/btnEXEC"


def test_find_by_table_cell():
    finder = SemanticFinder(FakeWindow())
    assert finder.find('col:"Status" row:1').endswith("col[Status]")


def test_suggest_returns_hints():
    finder = SemanticFinder(FakeWindow())
    hints = finder.suggest(
        ElementInfo("wnd[0]/usr/btnEXEC", "Execute", "Execute", "GuiButton", None)
    )
    assert any(h.locator_type is LocatorType.H_LABEL for h in hints)
