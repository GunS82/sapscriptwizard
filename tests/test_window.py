from __future__ import annotations

from typing import Any
from unittest.mock import MagicMock

import pytest

from sapscriptwizard.core.errors import ErrorCode, SapError
from sapscriptwizard.core.window import Window


class DummyElement:
    def __init__(self, text: str = "") -> None:
        self.Text = text
        self.Enabled = True
        self.Visible = True

    def Press(self) -> None:
        self.Text = "pressed"

    def SetFocus(self) -> None:  # pragma: no cover - stub
        pass

    def Text(self, value: str) -> None:  # pragma: no cover - legacy setter signature
        self.Text = value


class DummyWindow:
    def __init__(self) -> None:
        self.elements = {"wnd[0]/usr/txt1": DummyElement("value")}
        self.children: list[Any] = []
        self.Text = "Main"
        self.Id = "wnd[0]"

    def FindById(self, element_id: str) -> DummyElement:
        if element_id not in self.elements:
            raise RuntimeError("not found")
        return self.elements[element_id]

    def SendVKey(self, code: int) -> None:
        self.last_vkey = code

    @property
    def Children(self) -> MagicMock:
        collection = MagicMock()
        collection.Count = len(self.children)
        collection.Item.side_effect = self.children.__getitem__
        return collection


class DummyGateway:
    def call_method(self, obj, name: str, *args, **kwargs):
        return getattr(obj, name)(*args, **kwargs)

    def get_attr(self, obj, name: str):
        return getattr(obj, name)


def test_read_and_press():
    dummy_window = DummyWindow()
    window = Window(dummy_window, gateway=DummyGateway())
    assert window.read_text("wnd[0]/usr/txt1") == "value"
    window.press("wnd[0]/usr/txt1")
    assert dummy_window.elements["wnd[0]/usr/txt1"].Text == "pressed"


def test_find_missing_element():
    window = Window(DummyWindow(), gateway=DummyGateway())
    with pytest.raises(SapError) as excinfo:
        window.read_text("missing")
    assert excinfo.value.code is ErrorCode.NOT_FOUND


def test_dump_structure_includes_window():
    window = Window(DummyWindow(), gateway=DummyGateway())
    structure = window.dump_gui_structure(max_depth=0)
    assert structure["window_id"] == "wnd[0]"
    assert structure["title"] == "Main"
