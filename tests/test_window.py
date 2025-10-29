from __future__ import annotations

from dataclasses import dataclass

import pytest

from sapscriptwizard.core.session import SapSession
from sapscriptwizard.core.window import Window


@dataclass
class DummyElement:
    Id: str
    Type: str
    Text: str
    Value: str | None = None
    Enabled: bool = True
    Visible: bool = True
    Children: tuple["DummyElement", ...] = ()

    def press(self) -> None:
        self.pressed = True


class DummyStatusBar:
    def __init__(self) -> None:
        self.Text = "Ready"
        self.MessageType = "S"


class DummyWindow:
    def __init__(self) -> None:
        self.Text = "Main"
        self.StatusBar = DummyStatusBar()
        self.Children = (
            DummyElement("usr/txtField", "GuiTextField", "Field", "value"),
            DummyElement("usr/btnPress", "GuiButton", "Press"),
        )

    def sendVKey(self, code: int) -> None:  # noqa: N802 - COM naming
        self.last_vkey = code


class DummySession:
    def __init__(self) -> None:
        self.elements = {
            "wnd[0]": DummyWindow(),
            "wnd[0]/usr/txtField": DummyElement("usr/txtField", "GuiTextField", "Field", "value"),
            "wnd[0]/usr/btnPress": DummyElement("usr/btnPress", "GuiButton", "Press"),
        }

    def findById(self, element_id: str):  # noqa: N802 - COM naming
        if element_id not in self.elements:
            raise KeyError(element_id)
        return self.elements[element_id]


def test_window_read_write_and_press():
    session = SapSession(DummySession())
    window = Window(session, "wnd[0]")
    assert window.read_text("usr/txtField") == "value"
    window.write("usr/txtField", "new")
    assert session.raw.findById("wnd[0]/usr/txtField").Value == "new"
    window.press("usr/btnPress")
    assert session.raw.findById("wnd[0]/usr/btnPress").pressed is True


def test_window_dump_structure():
    session = SapSession(DummySession())
    window = Window(session, "wnd[0]")
    dump = window.dump_gui_structure(max_depth=2)
    assert dump["window_id"] == "wnd[0]"
    assert dump["title"] == "Main"
    assert any(element["id"].endswith("txtField") for element in dump["elements"])


def test_window_send_vkey():
    session = SapSession(DummySession())
    window = Window(session, "wnd[0]")
    window.send_vkey(0)
    assert session.raw.findById("wnd[0]").last_vkey == 0


def test_window_write_disabled():
    session = SapSession(DummySession())
    disabled = session.raw.findById("wnd[0]/usr/txtField")
    disabled.Enabled = False
    window = Window(session, "wnd[0]")
    from sapscriptwizard.core.errors import SapError, SapErrorCode

    with pytest.raises(SapError) as exc:
        window.write("usr/txtField", "x")
    assert exc.value.code is SapErrorCode.DISABLED
