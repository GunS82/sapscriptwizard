import math
from typing import Optional, List, Tuple
from dataclasses import dataclass, field
import win32com.client
import re
import logging

log = logging.getLogger(__name__)

@dataclass(frozen=True)
class Position:
    """Хранит позицию и размеры элемента на экране."""
    left: int
    top: int
    width: int
    height: int
    # --- Вычисляемые свойства ---
    right: int = field(init=False)
    bottom: int = field(init=False)
    center_x: int = field(init=False)
    center_y: int = field(init=False)

    def __post_init__(self):
        # Используем object.__setattr__, т.к. dataclass заморожен (frozen=True)
        object.__setattr__(self, 'right', self.left + self.width)
        object.__setattr__(self, 'bottom', self.top + self.height)
        object.__setattr__(self, 'center_x', self.left + self.width // 2)
        object.__setattr__(self, 'center_y', self.top + self.height // 2)

    def is_horizontally_aligned_with(self, other: 'Position', tolerance: int = 5) -> bool:
        """Проверяет, выровнены ли центры по вертикали (горизонтальное выравнивание)."""
        return abs(self.center_y - other.center_y) <= tolerance

    def is_vertically_aligned_with(self, other: 'Position', tolerance: int = 8) -> bool:
        """Проверяет, выровнены ли центры по горизонтали (вертикальное выравнивание)."""
        # Используем только center_x для простоты, можно добавить left/right align
        return abs(self.center_x - other.center_x) <= tolerance

    def is_right_of(self, other: 'Position', gap_tolerance: int = 25) -> bool:
        """Проверяет, находится ли этот элемент справа от другого с допустимым зазором."""
        distance = self.left - other.right
        return distance >= 0 and distance <= gap_tolerance

    def is_left_of(self, other: 'Position', gap_tolerance: int = 25) -> bool:
        """Проверяет, находится ли этот элемент слева от другого с допустимым зазором."""
        distance = other.left - self.right
        return distance >= 0 and distance <= gap_tolerance

    def is_below(self, other: 'Position', gap_tolerance: int = 25) -> bool:
        """Проверяет, находится ли этот элемент ниже другого с допустимым зазором."""
        distance = self.top - other.bottom
        return distance >= 0 and distance <= gap_tolerance

    def is_above(self, other: 'Position', gap_tolerance: int = 25) -> bool:
        """Проверяет, находится ли этот элемент выше другого с допустимым зазором."""
        distance = other.top - self.bottom
        return distance >= 0 and distance <= gap_tolerance

    def distance_squared_to(self, other: 'Position') -> int:
        """Возвращает квадрат расстояния между центрами элементов."""
        dx = self.center_x - other.center_x
        dy = self.center_y - other.center_y
        return dx*dx + dy*dy

@dataclass
class ElementInfo:
    """Хранит информацию о найденном GUI элементе."""
    element_id: str
    element_type: str # e.g., "GuiTextField", "GuiLabel"
    text: Optional[str]
    tooltip: Optional[str]
    position: Position
    name: Optional[str] = None # SAP Name property, if available
    changeable: Optional[bool] = None

# Типы локаторов для внутреннего использования
@dataclass
class LocatorStrategy: pass

@dataclass
class ContentLocator(LocatorStrategy): value: str

@dataclass
class HLabelLocator(LocatorStrategy): label: str

@dataclass
class VLabelLocator(LocatorStrategy): label: str

@dataclass
class HLabelVLabelLocator(LocatorStrategy): h_label: str; v_label: str

@dataclass
class HLabelHLabelLocator(LocatorStrategy): left_label: str; right_label: str
# Добавить HIndexVLabelLocator, HLabelVIndexLocator если нужно
# --- END OF FILE: pysapscript/locator_helpers.py ---