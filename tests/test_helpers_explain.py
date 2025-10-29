from __future__ import annotations

from sapscriptwizard.helpers.explain import explain_id


def test_explain_id_returns_element():
    structure = {
        "window_id": "wnd[0]",
        "title": "Main",
        "containers": [
            {
                "id": "child",
                "type": "GuiButton",
                "tech_name": "BTN1",
                "label": "Execute",
                "value": None,
                "enabled": True,
                "visible": True,
                "children": [],
            }
        ],
    }
    element = explain_id(structure, "child")
    assert element.element_id == "child"
    assert element.label == "Execute"
