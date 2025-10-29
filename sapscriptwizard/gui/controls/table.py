"""Utilities for working with ALV/Shell tables."""
from __future__ import annotations

from typing import Any, Iterable, Iterator, Mapping

from ...core.errors import SapError, SapErrorCode


class ShellTable:
    """Wrapper around SAP GuiGridView like controls."""

    def __init__(self, window: "Window", element_id: str) -> None:
        from ...core.window import Window

        if not isinstance(window, Window):  # pragma: no cover - defensive
            raise TypeError("window must be Window")
        self._window = window
        self._element_id = element_id
        self._raw = window.find_by_id(element_id)

    # ------------------------------------------------------------------
    @property
    def raw(self) -> Any:
        return self._raw

    def to_dicts(self, limit_rows: int | None = None) -> list[dict[str, Any]]:
        columns = self._read_columns()
        rows = self._read_rows(limit_rows)
        return [dict(zip(columns, row)) for row in rows]

    def _read_columns(self) -> list[str]:
        columns = getattr(self._raw, "ColumnOrder", None)
        if columns is None:
            raise SapError(SapErrorCode.UNSUPPORTED, "Table does not expose ColumnOrder")
        if isinstance(columns, (list, tuple)):
            return [str(col) for col in columns]
        if hasattr(columns, "Count") and hasattr(columns, "Item"):
            return [str(columns.Item(idx)) for idx in range(columns.Count)]
        if hasattr(columns, "__iter__"):
            return [str(item) for item in columns]
        raise SapError(SapErrorCode.UNSUPPORTED, "Unsupported ColumnOrder container")

    def _row_count(self) -> int:
        return int(getattr(self._raw, "RowCount", 0))

    def _read_rows(self, limit_rows: int | None) -> list[list[Any]]:
        row_count = self._row_count()
        if limit_rows is not None:
            row_count = min(row_count, limit_rows)
        data: list[list[Any]] = []
        for row_idx in range(row_count):
            data.append(self._read_row(row_idx))
        return data

    def _read_row(self, row_idx: int) -> list[Any]:
        columns = self._read_columns()
        values: list[Any] = []
        for column in columns:
            getter = getattr(self._raw, "GetCellValue", None)
            if not callable(getter):
                raise SapError(SapErrorCode.UNSUPPORTED, "Table does not support GetCellValue")
            values.append(getter(row_idx, column))
        return values


__all__ = ["ShellTable"]
