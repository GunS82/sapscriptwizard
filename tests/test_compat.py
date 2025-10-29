from __future__ import annotations

import warnings

from sapscriptwizard.compat.legacy import ShellTable, Window
from sapscriptwizard.core.window import Window as NewWindow


def test_legacy_window_wraps_new_window():
    class DummyWindow:
        def __init__(self) -> None:
            self.calls = []

        def FindById(self, element_id: str) -> DummyElement:
            return DummyElement(element_id)

    class DummyElement:
        def __init__(self, element_id: str) -> None:
            self.element_id = element_id
            self.Text = "text"

        def Press(self) -> None:
            self.Text = "pressed"

    new = NewWindow(DummyWindow())
    with warnings.catch_warnings(record=True) as caught:
        legacy = Window(new)
        legacy.press("id")
        assert any(isinstance(w.message, DeprecationWarning) for w in caught)


def test_legacy_shell_table_warns():
    class DummyTable:
        ColumnOrder = ["A"]
        RowCount = 1

        def GetCellValue(self, row: int, column: str) -> str:
            return "value"

    with warnings.catch_warnings(record=True) as caught:
        table = ShellTable(DummyTable())
        assert table.to_dicts() == [{"A": "value"}]
        assert any(isinstance(w.message, DeprecationWarning) for w in caught)
