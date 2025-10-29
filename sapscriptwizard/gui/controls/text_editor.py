"""Generic text editor helpers."""
from __future__ import annotations

from ...core.errors import SapError, SapErrorCode


class GuiTextEditor:
    def __init__(self, window: "Window", element_id: str) -> None:
        from ...core.window import Window

        if not isinstance(window, Window):  # pragma: no cover - defensive
            raise TypeError("window must be Window")
        self._window = window
        self._element_id = element_id
        self._raw = window.find_by_id(element_id)

    def get_text(self) -> str:
        getter = getattr(self._raw, "Text", None)
        if isinstance(getter, str):
            return getter
        if callable(getter):
            return getter()
        raise SapError(SapErrorCode.UNSUPPORTED, "Text editor does not expose text getter")

    def set_text(self, text: str) -> None:
        setter = getattr(self._raw, "Text", None)
        if callable(setter):
            self._window.gateway.call(setter, text)
            return
        if hasattr(self._raw, "Text"):
            setattr(self._raw, "Text", text)
            return
        raise SapError(SapErrorCode.UNSUPPORTED, "Text editor does not expose text setter")


__all__ = ["GuiTextEditor"]
