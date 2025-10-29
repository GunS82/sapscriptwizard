from __future__ import annotations

from sapscriptwizard.helpers.explain import explain_id, suggest_locator


class FakeWindow:
    def __init__(self) -> None:
        self._structure = {
            "containers": [],
            "elements": [
                {
                    "id": "wnd[0]/usr/btnEXEC",
                    "type": "GuiButton",
                    "label": "Execute",
                    "value": "Execute",
                    "tech_name": "BTN_EXEC",
                    "children": [],
                }
            ],
        }

    def dump_gui_structure(self, max_depth: int = 2, max_children: int = 100):
        return self._structure

    @property
    def finder(self):
        from sapscriptwizard.gui.finders.semantic_finder import SemanticFinder

        return SemanticFinder(self)


def test_explain_and_suggest():
    window = FakeWindow()
    info = explain_id(window, "wnd[0]/usr/btnEXEC")
    assert info.label == "Execute"
    hints = suggest_locator(window, "wnd[0]/usr/btnEXEC")
    assert hints
