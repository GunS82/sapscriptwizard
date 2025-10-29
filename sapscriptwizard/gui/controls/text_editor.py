"""Generic GuiTextEdit helper."""
from __future__ import annotations

from typing import Any

from ...core.com_gateway import ComGateway
from ...core.errors import ErrorCode, SapError


class GuiTextEditor:
    def __init__(self, com_editor: Any, *, gateway: ComGateway | None = None) -> None:
        self._editor = com_editor
        self.gateway = gateway or ComGateway()

    def get_all_text(self) -> str:
        try:
            return str(self.gateway.get_attr(self._editor, "Text"))
        except SapError:
            raise
        except Exception as exc:
            raise SapError(
                ErrorCode.UNSUPPORTED,
                "Text editor has no Text attribute",
                str(exc),
            ) from exc

    def set_all_text(self, text: str) -> None:
        try:
            self.gateway.invoke(lambda: setattr(self._editor, "Text", text))
        except SapError:
            raise
        except Exception as exc:
            raise SapError(ErrorCode.DISABLED, "Failed to write text editor", str(exc)) from exc
