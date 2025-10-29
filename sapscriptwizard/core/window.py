"""High level window operations."""
from __future__ import annotations

import logging
from typing import Any

from .com_gateway import ComGateway
from .errors import ErrorCode, SapError
from .types import ElementInfo

log = logging.getLogger(__name__)


class Window:
    """Wrapper around a SAP GUI window COM object."""

    def __init__(self, com_window: Any, *, gateway: ComGateway | None = None) -> None:
        self._com_window = com_window
        self.gateway = gateway or ComGateway()

    def _ensure_element(self, element_id: str) -> Any:
        try:
            return self.gateway.call_method(self._com_window, "FindById", element_id)
        except SapError:
            raise
        except Exception as exc:
            raise SapError(
                ErrorCode.NOT_FOUND,
                f"Element {element_id} not found",
                str(exc),
            ) from exc

    def find_by_id(self, element_id: str) -> Any:
        """Return the raw COM element for *element_id*."""

        return self._ensure_element(element_id)

    def read_text(self, element_id: str) -> str:
        element = self._ensure_element(element_id)
        try:
            return str(self.gateway.get_attr(element, "Text"))
        except SapError:
            raise
        except Exception as exc:
            raise SapError(ErrorCode.UNSUPPORTED, "Element has no text", str(exc)) from exc

    def write(self, element_id: str, text: str) -> None:
        element = self._ensure_element(element_id)
        try:
            self.gateway.call_method(element, "SetFocus")
        except SapError:
            pass
        except Exception as exc:
            log.debug("SetFocus failed for %s: %s", element_id, exc)

        try:
            self.gateway.invoke(lambda: setattr(element, "Text", text))
        except SapError:
            raise
        except Exception as exc:
            raise SapError(ErrorCode.DISABLED, "Failed writing to element", str(exc)) from exc

    def press(self, element_id: str) -> None:
        element = self._ensure_element(element_id)
        try:
            self.gateway.call_method(element, "Press")
        except SapError:
            raise
        except Exception as exc:
            raise SapError(ErrorCode.DISABLED, "Element cannot be pressed", str(exc)) from exc

    def send_vkey(self, code: int) -> None:
        try:
            self.gateway.call_method(self._com_window, "SendVKey", code)
        except SapError:
            raise
        except Exception as exc:
            raise SapError(ErrorCode.UNSUPPORTED, "Failed to send VKey", str(exc)) from exc

    def dump_gui_structure(self, max_depth: int = 2, max_children: int = 100) -> dict[str, Any]:
        """Return a dictionary describing the GUI structure."""

        def make_element(com_obj: Any, depth: int) -> ElementInfo:
            element_id = getattr(com_obj, "Id", "")
            info = ElementInfo(
                element_id=element_id,
                type=getattr(com_obj, "Type", ""),
                tech_name=getattr(com_obj, "Name", None),
                label=getattr(com_obj, "Text", None),
                value=getattr(com_obj, "Value", None),
                enabled=getattr(com_obj, "Enabled", None),
                visible=getattr(com_obj, "Visible", None),
            )

            if depth >= max_depth:
                return info

            try:
                children = com_obj.Children
            except AttributeError:
                return info

            if children is None:
                return info

            try:
                count = min(getattr(children, "Count", 0), max_children)
            except Exception:
                count = 0

            for idx in range(count):
                try:
                    child = self.gateway.call_method(children, "Item", idx)
                except SapError:
                    log.debug("Unable to obtain child %s for %s", idx, element_id, exc_info=True)
                    continue
                info.children.append(make_element(child, depth + 1))

            return info

        root_info = make_element(self._com_window, 0)
        containers: list[dict[str, Any]] = []
        elements: list[dict[str, Any]] = []
        for child in root_info.children:
            if child.children:
                containers.append(child.to_dict())
            else:
                elements.append(child.to_dict())

        return {
            "window_id": getattr(self._com_window, "Id", ""),
            "title": getattr(self._com_window, "Text", ""),
            "status_bar": getattr(getattr(self._com_window, "StatusBar", None), "Text", ""),
            "containers": containers,
            "elements": elements,
        }
