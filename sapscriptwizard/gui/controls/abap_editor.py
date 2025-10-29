"""ABAP editor helpers."""
from __future__ import annotations

from ...core.errors import SapError, SapErrorCode


class GuiAbapEditor:
    def __init__(self, window: "Window", element_id: str) -> None:
        from ...core.window import Window

        if not isinstance(window, Window):  # pragma: no cover - defensive
            raise TypeError("window must be Window")
        self._window = window
        self._element_id = element_id
        self._raw = window.find_by_id(element_id)

    def get_all_text(self) -> str:
        getter = getattr(self._raw, "Text", None)
        if isinstance(getter, str):
            return getter
        if callable(getter):
            return getter()
        get_method = getattr(self._raw, "GetText", None)
        if callable(get_method):
            return get_method()
        raise SapError(SapErrorCode.UNSUPPORTED, "Editor does not expose text getter")

    def set_text(self, text: str) -> None:
        setter = getattr(self._raw, "Text", None)
        if not callable(setter) and not isinstance(setter, str):
            setter = getattr(self._raw, "SetText", None)
        if callable(setter):
            self._window.gateway.call(setter, text)
            return
        if hasattr(self._raw, "Text"):
            setattr(self._raw, "Text", text)
            return
        raise SapError(SapErrorCode.UNSUPPORTED, "Editor does not expose text setter")


__all__ = ["GuiAbapEditor"]
