from __future__ import annotations

from sapscriptwizard.core.session import SapSession
from sapscriptwizard.gui.controls.table import ShellTable


class DummyTable:
    def __init__(self) -> None:
        self.ColumnOrder = ["Status", "Message"]
        self.RowCount = 2
        self._data = [
            {"Status": "OK", "Message": "All good"},
            {"Status": "WARN", "Message": "Check"},
        ]

    def GetCellValue(self, row: int, column: str):  # noqa: N802 - COM naming
        return self._data[row][column]


class DummySession:
    def __init__(self) -> None:
        self.elements = {
            "wnd[0]": self,
            "wnd[0]/usr/cntlGRID": DummyTable(),
        }

    def findById(self, element_id: str):  # noqa: N802 - COM naming
        return self.elements[element_id]


def test_shell_table_to_dicts():
    session = SapSession(DummySession())
    table = ShellTable(session.window(), "usr/cntlGRID")
    data = table.to_dicts()
    assert data[0]["Status"] == "OK"
    assert len(data) == 2
