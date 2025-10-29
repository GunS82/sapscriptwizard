"""Abstractions for GuiShell table controls."""
from __future__ import annotations

from collections.abc import Iterable, Sequence
from typing import Any

from ...core.com_gateway import ComGateway
from ...core.errors import ErrorCode, SapError


class ShellTable:
    """Light-weight reader for GuiGridView tables."""

    def __init__(self, com_table: Any, *, gateway: ComGateway | None = None) -> None:
        self._table = com_table
        self.gateway = gateway or ComGateway()
        self.columns = list(self._resolve_columns())

    def _resolve_columns(self) -> Iterable[str]:
        order = getattr(self._table, "ColumnOrder", [])
        if isinstance(order, Sequence):
            return [str(col) for col in order]
        count = getattr(order, "Count", 0)
        return [str(self.gateway.call_method(order, "Item", idx)) for idx in range(count)]

    def _row_count(self) -> int:
        try:
            return int(self._table.RowCount)
        except Exception as exc:
            raise SapError(
                ErrorCode.UNSUPPORTED,
                "Table does not expose RowCount",
                str(exc),
            ) from exc

    def _cell_value(self, row: int, column: str) -> Any:
        try:
            return self.gateway.call_method(self._table, "GetCellValue", row, column)
        except SapError:
            raise
        except Exception as exc:
            raise SapError(ErrorCode.COM_FAIL, "Could not read table cell", str(exc)) from exc

    def to_dicts(self, limit_rows: int | None = None) -> list[dict[str, Any]]:
        """Return the table data as a list of dictionaries."""

        result: list[dict[str, Any]] = []
        row_count = self._row_count()
        max_row = min(row_count, limit_rows) if limit_rows is not None else row_count
        for row in range(max_row):
            payload = {column: self._cell_value(row, column) for column in self.columns}
            result.append(payload)
        return result
