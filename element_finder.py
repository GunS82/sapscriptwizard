# --- START OF FILE: pysapscript/element_finder.py ---
# (This is a new file)
import win32com.client
import logging
import time
import re
from typing import Optional, List, Dict, Any

# Используем относительный импорт для хелперов
from .locator_helpers import (
    Position, ElementInfo, LocatorStrategy, ContentLocator, HLabelLocator,
    VLabelLocator, HLabelVLabelLocator, HLabelHLabelLocator
)
# Используем относительный импорт для исключений
from .types_ import exceptions

log = logging.getLogger(__name__)

# Список типов элементов, которые обычно являются целью локаторов (полей ввода, кнопок и т.д.)
DEFAULT_TARGET_TYPES = [
    "GuiTextField", "GuiCTextField", "GuiPasswordField",
    "GuiComboBox",
    "GuiCheckBox",
    "GuiRadioButton",
    "GuiButton",
    "GuiTab", # Иногда метка может указывать на вкладку
]
# Типы элементов, которые могут выступать в роли меток
LABEL_ELEMENT_TYPES = ["GuiLabel", "GuiTextField", "GuiCTextField"] # Иногда поля без возможности ввода используются как заголовки/метки

class SapElementFinder:
    """
    Находит элементы SAP GUI, используя семантические локаторы (по меткам, содержимому, позиции).
    """
    def __init__(self, session_handle: win32com.client.CDispatch):
        self.session_handle = session_handle
        self._element_cache: Dict[str, List[ElementInfo]] = {} # Кэш элементов {type: [ElementInfo]}
        self._cache_window_id: Optional[str] = None # ID окна, для которого актуален кэш

    def _check_and_refresh_cache(self) -> None:
        """Проверяет, изменилось ли активное окно, и обновляет кэш, если нужно."""
        try:
            current_window = self.session_handle.ActiveWindow
            current_window_id = current_window.Id
            if current_window_id != self._cache_window_id:
                log.info(f"Window changed (from '{self._cache_window_id}' to '{current_window_id}'). Refreshing element cache.")
                self._scan_window_elements(current_window)
                self._cache_window_id = current_window_id
            # else: log.debug("Cache is up to date.") # Можно раскомментировать для отладки
        except Exception as e:
            # Ошибка при доступе к ActiveWindow может означать, что сессия не активна
            log.error(f"Failed to check/refresh element cache: {e}. Clearing cache.")
            self._clear_cache()
            # Перебрасываем исключение, т.к. без окна работать нельзя
            raise exceptions.SapGuiComException(f"Error accessing ActiveWindow: {e}") from e

    def _clear_cache(self) -> None:
        """Очищает кэш элементов."""
        self._element_cache = {}
        self._cache_window_id = None
        log.debug("Element cache cleared.")

    def _scan_window_elements(self, root_element: win32com.client.CDispatch) -> None:
        """Сканирует элементы окна и заполняет кэш."""
        self._clear_cache()
        log.debug(f"Scanning elements starting from '{root_element.Id}'...")
        start_time = time.time()
        elements_found = 0

        queue = [root_element]
        processed_ids = set() # Для предотвращения бесконечных циклов на некоторых структурах

        while queue:
            component = queue.pop(0)

            try:
                 component_id = component.Id
                 if component_id in processed_ids:
                     continue
                 processed_ids.add(component_id)

                 # --- Получаем базовую информацию ---
                 elem_type = getattr(component, "Type", "Unknown")
                 elem_text: Optional[str] = None
                 elem_tooltip: Optional[str] = None
                 elem_name: Optional[str] = None
                 elem_changeable: Optional[bool] = None
                 elem_pos: Optional[Position] = None

                 # Пытаемся получить основные свойства безопасно
                 try: elem_text = str(getattr(component, "Text", "")).strip()
                 except: pass
                 try: elem_tooltip = str(getattr(component, "Tooltip", "")).strip()
                 except: pass # Tooltip иногда вызывает ошибки
                 try: elem_name = str(getattr(component, "Name", "")).strip()
                 except: pass
                 try: elem_changeable = bool(getattr(component, "Changeable", False))
                 except: pass
                 try:
                     # Убедимся, что все координаты доступны
                     if all(hasattr(component, attr) for attr in ["ScreenLeft", "ScreenTop", "Width", "Height"]):
                         elem_pos = Position(
                             left=component.ScreenLeft, top=component.ScreenTop,
                             width=component.Width, height=component.Height
                         )
                 except: pass # Ошибка получения координат

                 # --- Если есть позиция, добавляем в кэш ---
                 if elem_pos:
                     info = ElementInfo(
                         element_id=component_id,
                         element_type=elem_type,
                         text=elem_text,
                         tooltip=elem_tooltip,
                         position=elem_pos,
                         name=elem_name,
                         changeable=elem_changeable
                     )
                     if elem_type not in self._element_cache:
                         self._element_cache[elem_type] = []
                     self._element_cache[elem_type].append(info)
                     elements_found += 1

                 # --- Добавляем дочерние элементы в очередь ---
                 if getattr(component, "ContainerType", False) and hasattr(component, "Children"):
                     try:
                         children = component.Children
                         child_count = getattr(children, "Count", 0)
                         for i in range(child_count):
                             try:
                                 child = children(i) # Доступ к элементу коллекции
                                 if child: queue.append(child)
                             except Exception as child_e:
                                  # Логируем ошибку доступа к конкретному дочернему элементу, но продолжаем
                                  log.warning(f"Could not access child at index {i} of {component_id}: {child_e}")
                     except Exception as children_e:
                          # Логируем ошибку доступа к коллекции Children, но продолжаем
                          log.warning(f"Could not access children of {component_id}: {children_e}")

            except Exception as component_e:
                 # Логируем ошибку обработки элемента, но продолжаем сканирование
                 log.warning(f"Error processing component during scan: {component_e}")

        end_time = time.time()
        log.info(f"Element scan complete. Found {elements_found} elements with positions in {end_time - start_time:.3f} seconds.")
        # log.debug(f"Cache content: {self._element_cache}") # Отладка: показать кэш

    def _parse_locator(self, locator_str: str) -> LocatorStrategy:
        """Парсит строку локатора и возвращает объект стратегии."""
        locator_str = locator_str.strip()

        # 1. Content Locator (=)
        if locator_str.startswith('='):
            return ContentLocator(value=locator_str[1:].strip())

        # 2. HLabelHLabel (>>)
        if '>>' in locator_str:
            parts = [p.strip() for p in locator_str.split('>>', 1)]
            if len(parts) == 2 and parts[0] and parts[1]:
                return HLabelHLabelLocator(left_label=parts[0], right_label=parts[1])
            else:
                raise ValueError(f"Invalid HLabelHLabel locator format: '{locator_str}'")

        # 3. Locators with '@' (VLabel, HLabelVLabel, HIndexVLabel, HLabelVIndex)
        if '@' in locator_str:
            parts = [p.strip() for p in locator_str.split('@', 1)]
            left, right = parts[0], parts[1]

            # 3a. VLabel (@ label)
            if not left and right:
                return VLabelLocator(label=right)

            # 3b. HLabelVLabel (label @ label) / HIndexVLabel / HLabelVIndex
            if left and right:
                 # Пытаемся определить, индекс ли слева или справа
                 left_is_index = left.isdigit()
                 right_is_index = right.isdigit()

                 if left_is_index and not right_is_index:
                      # return HIndexVLabelLocator(h_index=int(left), v_label=right) # Пока не реализовано
                      raise NotImplementedError("HIndexVLabelLocator (index @ label) is not implemented yet.")
                 elif not left_is_index and right_is_index:
                      # return HLabelVIndexLocator(h_label=left, v_index=int(right)) # Пока не реализовано
                      raise NotImplementedError("HLabelVIndexLocator (label @ index) is not implemented yet.")
                 elif not left_is_index and not right_is_index:
                      # Убираем кавычки, если метка была числом в кавычках
                      h_label = left.strip('"')
                      v_label = right.strip('"')
                      return HLabelVLabelLocator(h_label=h_label, v_label=v_label)
                 else: # Оба - числа, некорректный формат
                      raise ValueError(f"Invalid locator format with '@': '{locator_str}'")
            else: # Одна из частей пуста (кроме случая @ label)
                 raise ValueError(f"Invalid locator format with '@': '{locator_str}'")

        # 4. HLabel (простая метка)
        if locator_str:
            return HLabelLocator(label=locator_str)

        # 5. Некорректный локатор
        raise ValueError(f"Could not parse locator: '{locator_str}'")

    def _find_label_element(self, label_text: str) -> Optional[ElementInfo]:
        """Находит элемент-метку по тексту (в кэше)."""
        all_label_candidates: List[ElementInfo] = []
        for label_type in LABEL_ELEMENT_TYPES:
             all_label_candidates.extend(self._element_cache.get(label_type, []))

        for label in all_label_candidates:
            if label.text == label_text:
                return label
        return None

    def _filter_by_type(self, elements: List[ElementInfo], target_types: Optional[List[str]]) -> List[ElementInfo]:
         """Фильтрует список элементов по заданным типам."""
         if target_types is None:
             return elements
         return [elem for elem in elements if elem.element_type in target_types]

    def find_element(self, locator_str: str, target_element_types: Optional[List[str]] = None) -> Optional[str]:
        """
        Основной метод поиска элемента по семантическому локатору.

        Args:
            locator_str: Строка локатора (e.g., "Пользователь", "@ Пароль", "=Сохранить").
            target_element_types: Опциональный список типов элементов для поиска
                                  (e.g., ["GuiTextField", "GuiCTextField"]). Если None,
                                  используется DEFAULT_TARGET_TYPES.

        Returns:
            ID найденного элемента или None.
        """
        self._check_and_refresh_cache()
        if not self._element_cache:
            log.warning("Element cache is empty. Cannot find element.")
            return None

        try:
            strategy = self._parse_locator(locator_str)
            log.debug(f"Parsed locator '{locator_str}' as: {strategy}")
        except (ValueError, NotImplementedError) as e:
            log.error(f"Error parsing locator '{locator_str}': {e}")
            return None

        # Определяем целевые типы
        effective_target_types = target_element_types if target_element_types is not None else DEFAULT_TARGET_TYPES

        # Собираем все потенциально целевые элементы из кэша
        candidate_elements: List[ElementInfo] = []
        for elem_type in effective_target_types:
             candidate_elements.extend(self._element_cache.get(elem_type, []))

        if not candidate_elements:
             log.warning(f"No candidate elements found for types: {effective_target_types}")
             return None

        found_element: Optional[ElementInfo] = None

        # --- Реализация стратегий поиска ---

        if isinstance(strategy, ContentLocator):
            # Ищем по тексту или тултипу среди ВСЕХ кэшированных элементов (не только target_types)
            all_elements = [item for sublist in self._element_cache.values() for item in sublist]
            for elem in all_elements:
                # Приоритет тексту, потом тултипу
                if elem.text == strategy.value:
                    found_element = elem
                    break
                if elem.tooltip == strategy.value:
                    found_element = elem
                    break # Нашли по тултипу, выходим

        elif isinstance(strategy, HLabelLocator):
            label_elem = self._find_label_element(strategy.label)
            if label_elem:
                # Ищем ближайший справа и горизонтально выровненный
                closest_right: Optional[ElementInfo] = None
                min_dist = float('inf')
                for target in candidate_elements:
                    if target.position.is_horizontally_aligned_with(label_elem.position) and \
                       target.position.is_right_of(label_elem.position):
                        dist = target.position.left - label_elem.position.right
                        if dist < min_dist:
                            min_dist = dist
                            closest_right = target
                found_element = closest_right
            else:
                log.debug(f"Label '{strategy.label}' not found for HLabel search.")

        elif isinstance(strategy, VLabelLocator):
            label_elem = self._find_label_element(strategy.label)
            if label_elem:
                # Ищем ближайший снизу и вертикально выровненный
                closest_below: Optional[ElementInfo] = None
                min_dist = float('inf')
                for target in candidate_elements:
                    if target.position.is_vertically_aligned_with(label_elem.position) and \
                       target.position.is_below(label_elem.position):
                        dist = target.position.top - label_elem.position.bottom
                        if dist < min_dist:
                            min_dist = dist
                            closest_below = target
                found_element = closest_below
            else:
                log.debug(f"Label '{strategy.label}' not found for VLabel search.")

        elif isinstance(strategy, HLabelVLabelLocator):
            h_label_elem = self._find_label_element(strategy.h_label)
            v_label_elem = self._find_label_element(strategy.v_label)
            if h_label_elem and v_label_elem:
                # Ищем элемент, который выровнен по горизонтали с h_label И по вертикали с v_label
                # и находится правее h_label и ниже v_label
                best_match: Optional[ElementInfo] = None
                min_dist_sq = float('inf')
                # Приблизительная точка пересечения (правый край h_label, нижний край v_label)
                cross_pos = Position(left=h_label_elem.position.right, top=v_label_elem.position.bottom, width=1, height=1)

                for target in candidate_elements:
                    if target.position.is_horizontally_aligned_with(h_label_elem.position) and \
                       target.position.is_vertically_aligned_with(v_label_elem.position) and \
                       target.position.left >= h_label_elem.position.right and \
                       target.position.top >= v_label_elem.position.bottom:
                        # Считаем расстояние до "точки пересечения"
                        dist_sq = target.position.distance_squared_to(cross_pos)
                        if dist_sq < min_dist_sq:
                            min_dist_sq = dist_sq
                            best_match = target
                found_element = best_match
            else:
                 log.debug(f"One or both labels not found for HLabelVLabel search: H='{strategy.h_label}', V='{strategy.v_label}'")

        elif isinstance(strategy, HLabelHLabelLocator):
             # Находим левый элемент (может быть меткой или полем)
             left_elem = self._find_label_element(strategy.left_label) or \
                         next((el for el in candidate_elements if el.text == strategy.left_label), None)
             if left_elem:
                  # Ищем правый элемент по тексту/тултипу среди кандидатов
                  possible_right_elements = [
                      el for el in candidate_elements
                      if (el.text == strategy.right_label or el.tooltip == strategy.right_label)
                         and el.position.is_horizontally_aligned_with(left_elem.position)
                         and el.position.is_right_of(left_elem.position)
                  ]
                  # Выбираем ближайший из найденных справа
                  if possible_right_elements:
                      found_element = min(possible_right_elements, key=lambda el: el.position.left - left_elem.position.right)
             else:
                  log.debug(f"Left element '{strategy.left_label}' not found for HLabelHLabel search.")


        # ... добавить реализацию для HIndexVLabelLocator и HLabelVIndexLocator ...

        # --- Результат ---
        if found_element:
            log.info(f"Locator '{locator_str}' resolved to element ID: {found_element.element_id} (Type: {found_element.element_type})")
            return found_element.element_id
        else:
            log.warning(f"Could not find element using locator: '{locator_str}'")
            return None

# --- END OF FILE: pysapscript/element_finder.py ---