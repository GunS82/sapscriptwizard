"""High level wrapper around a SAP GUI window."""
import pprint # Для dump_element_state
import json
import time
from time import sleep
from pathlib import Path
import re
from typing import Tuple, Optional, List, Generator, Any, Union, Dict, Pattern # Добавлено Dict, Optional, Union
import logging
import win32com.client

from sapscriptwizard.types_ import exceptions
from sapscriptwizard.types_.types import NavigateAction
from sapscriptwizard.shell_table import ShellTable
from sapscriptwizard.gui_tree import GuiTree
try:
    import yaml
except ImportError:
    yaml = None
try:
    from sapscriptwizard_semantic.element_finder import SapElementFinder, DEFAULT_TARGET_TYPES
except Exception:  # plugin not installed
    SapElementFinder = None  # type: ignore
    DEFAULT_TARGET_TYPES: List[str] = []


log = logging.getLogger(__name__)

class Window:
    def __init__(
        self,
        application: win32com.client.CDispatch,
        connection: int,
        connection_handle: win32com.client.CDispatch,
        session: int,
        session_handle: win32com.client.CDispatch,
        element_finder: Optional[Any] = None,
    ) -> None:
        self.application = application
        self.connection = connection
        self.connection_handle = connection_handle
        self.session = session
        self.session_handle = session_handle
        self._finder = element_finder

    def __repr__(self) -> str:
        return f"Window(connection={self.connection}, session={self.session})"

    def __str__(self) -> str:
        return f"Window(connection={self.connection}, session={self.session})"

    def __eq__(self, other: object) -> bool:
        if isinstance(other, Window):
            return self.connection == other.connection and self.session == other.session

        return False

    def __hash__(self) -> hash:
        return hash(f"{self.connection_handle}{self.session_handle}")

    # ... (существующие методы maximize, restore, close_window, navigate, etc.) ...
    # Копируем существующие методы для контекста
    def maximize(self) -> None:
        """ Maximizes this sap window """
        self.session_handle.findById("wnd[0]").maximize()

    def restore(self) -> None:
        """ Restores sap window to its default size, resp. before maximization """
        self.session_handle.findById("wnd[0]").restore()

    def close_window(self) -> None:
        """ Closes this sap window """
        self.session_handle.findById("wnd[0]").close()

    def navigate(self, action: NavigateAction) -> None:
        """ Navigates SAP: enter, back, end, cancel, save """
        el_map = {
            NavigateAction.enter: "wnd[0]/tbar[0]/btn[0]",
            NavigateAction.back: "wnd[0]/tbar[0]/btn[3]",
            NavigateAction.end: "wnd[0]/tbar[0]/btn[15]",
            NavigateAction.cancel: "wnd[0]/tbar[0]/btn[12]",
            NavigateAction.save: "wnd[0]/tbar[0]/btn[11]", # Corrected VKey for Save
        }
        el = el_map.get(action)
        if not el:
            raise exceptions.ActionException("Wrong navigation action!")
        self.press(el) # Use self.press for consistency and error handling

    def start_transaction(self, transaction: str) -> None:
        """ Starts transaction """
        self.write("wnd[0]/tbar[0]/okcd", transaction)
        self.press("wnd[0]/tbar[0]/btn[0]") # Send Enter

    def press(self, element: str) -> None:
        """ Presses element """
        try:
            self.session_handle.findById(element).press()
        except Exception as ex:
            raise exceptions.ActionException(f"Error pressing element {element}: {ex}")

    def select(self, element: str) -> None:
        """ Selects element or menu item """
        try:
            self.session_handle.findById(element).select()
        except Exception as ex:
            raise exceptions.ActionException(f"Error selecting element {element}: {ex}")

    def is_selected(self, element: str) -> bool:
        """ Gets status of select element """
        try:
            return self.session_handle.findById(element).selected
        except Exception as ex:
            raise exceptions.ActionException(f"Error getting status of element {element}: {ex}")

    def set_checkbox(self, element: str, selected: bool) -> None:
        """ Selects checkbox element """
        try:
            # Ensure the value passed is boolean
            checkbox = self.session_handle.findById(element)
            if checkbox.Type != "GuiCheckBox":
                 raise exceptions.ActionException(f"Element {element} is not a GuiCheckBox (Type: {checkbox.Type}).")
            checkbox.selected = bool(selected)
        except Exception as ex:
            raise exceptions.ActionException(f"Error setting checkbox {element}: {ex}")

    def write(self, element: str, text: str) -> None:
        """ Sets text property of an element """
        try:
            target_element = self.session_handle.findById(element)
            # Basic check if element likely accepts text
            if not hasattr(target_element, "text"):
                 raise exceptions.ActionException(f"Element {element} (Type: {target_element.Type}) does not seem to have a 'text' property.")
            target_element.text = str(text) # Ensure text is string
        except Exception as ex:
            raise exceptions.ActionException(f"Error writing to element {element}: {ex}")

    def read(self, element: str) -> str:
        """ Reads text property """
        try:
            target_element = self.session_handle.findById(element)
            if not hasattr(target_element, "text"):
                 raise exceptions.ActionException(f"Element {element} (Type: {target_element.Type}) does not seem to have a 'text' property.")
            return target_element.text
        except Exception as e:
            raise exceptions.ActionException(f"Error reading element {element}: {e}")

    def visualize(self, element: str, seconds: int = 1) -> None:
        """ draws red frame around the element """
        try:
            self.session_handle.findById(element).Visualize(True) # Use True instead of 1
            sleep(seconds)
            # Optional: Turn off visualization afterwards? SAP GUI might do this automatically.
            # self.session_handle.findById(element).Visualize(False)
        except Exception as e:
            raise exceptions.ActionException(f"Error visualizing element {element}: {e}")

    def exists(self, element: str) -> bool:
        """ checks if element exists by trying to access it """
        try:
            self.session_handle.findById(element)
            return True
        except Exception:
            return False

    def send_v_key(
        self,
        element: str = "wnd[0]",
        *,
        focus_element: Optional[str] = None, # Use Optional
        value: int = 0,
    ) -> None:
        """ Sends VKey to the window or element """
        try:
            target_element = self.session_handle.findById(element)
            if focus_element is not None:
                self.session_handle.findById(focus_element).SetFocus()
                sleep(0.1) # Small delay after setting focus might help
            target_element.sendVKey(value)
        except Exception as e:
            raise exceptions.ActionException(
                f"Error sending VKey {value} to element {element}"
                f"{' after focusing ' + focus_element if focus_element else ''}: {e}"
            )

    def read_html_viewer(self, element: str) -> str:
        """ Read the HTML content of the specified HTMLViewer element. """
        try:
            html_viewer = self.session_handle.findById(element)
            if html_viewer.Type != "GuiHTMLViewer":
                 raise exceptions.ActionException(f"Element {element} is not a GuiHTMLViewer.")
            # Accessing BrowserHandle might fail if content isn't fully loaded or control is different
            browser_handle = html_viewer.BrowserHandle
            # Wait briefly for document availability if necessary (experimental)
            # for _ in range(5):
            #     if browser_handle.Document: break
            #     sleep(0.2)
            if not browser_handle or not browser_handle.Document:
                 raise exceptions.ActionException(f"Could not access BrowserHandle or Document for {element}.")
            return browser_handle.Document.documentElement.innerHTML
        except Exception as e:
            raise exceptions.ActionException(f"Error reading HTMLViewer element {element}: {e}")

    def read_shell_table(self, element: str, load_table: bool = True) -> ShellTable:
        """ Read the table of the specified ShellTable/GuiGridView element. """
        return ShellTable(self.session_handle, element, load_table)

    # --- Существующие методы из Ver2 ---
    def get_status_message(self, window_id: str = "wnd[0]") -> Optional[Tuple[str, str, str, str]]:
        """ Reads the message from the status bar of the specified window. """
        statusbar_id = f"{window_id}/sbar"
        try:
            statusbar = self.session_handle.findById(statusbar_id)
            msg_text = getattr(statusbar, "Text", "")
            if msg_text:
                msg_type = getattr(statusbar, "MessageType", "")
                msg_id = getattr(statusbar, "MessageId", "").strip()
                msg_number = getattr(statusbar, "MessageNumber", "") # Often not set, but good to have
                return msg_type, msg_id, msg_number, msg_text
            return None
        except Exception as e:
            if "findById" in str(e):
                 return None
            else:
                 raise exceptions.StatusBarException(f"Error reading status bar '{statusbar_id}': {e}")
                
    def assert_status_bar(self,
                          window_id: str = "wnd[0]",
                          # --- Ожидаемые значения (любое или комбинация) ---
                          expected_type: Optional[Union[str, List[str]]] = None,
                          expected_id: Optional[Union[str, List[str]]] = None,
                          expected_number: Optional[Union[str, int, List[Union[str, int]]]] = None,
                          expected_text: Optional[Union[str, Pattern]] = None,
                          # --- Параметры ожидания и поведения ---
                          timeout: float = 2.0, # Время ожидания сообщения в секундах
                          poll_interval: float = 0.2, # Интервал проверки
                          raise_exception: bool = True,
                          fail_on_timeout: bool = True # Считать ли ошибкой, если сообщение не появилось
                         ) -> bool:
        """
        Проверяет соответствие сообщения в строке статуса заданным критериям.
        Пытается прочитать статус-бар несколько раз в течение указанного таймаута.

        Args:
            window_id (str): ID окна, в котором находится строка статуса.
            expected_type (Optional[Union[str, List[str]]]): Ожидаемый тип сообщения ('S', 'E', 'W', 'I', 'A')
                                                             или список допустимых типов.
            expected_id (Optional[Union[str, List[str]]]): Ожидаемый ID сообщения (напр., 'V1', '00')
                                                            или список допустимых ID.
            expected_number (Optional[Union[str, int, List[Union[str, int]]]]): Ожидаемый номер сообщения
                                                                               (строка или число) или список допустимых номеров.
                                                                               Преобразуется в строку для сравнения.
            expected_text (Optional[Union[str, Pattern]]): Ожидаемый текст сообщения. Может быть строкой для точного
                                                          совпадения (после strip) или скомпилированным
                                                          регулярным выражением (re.Pattern).
            timeout (float): Максимальное время ожидания появления сообщения (в секундах).
            poll_interval (float): Интервал между попытками чтения статус-бара (в секундах).
            raise_exception (bool): Выбрасывать ли StatusBarAssertionError, если проверка не пройдена.
                                    Если False, возвращает True/False.
            fail_on_timeout (bool): Считать ли ошибкой, если за время timeout сообщение так и не появилось
                                     в статус-баре (даже если критерии не заданы).

        Returns:
            bool: True, если сообщение соответствует всем заданным критериям.
                  False, если не соответствует (и raise_exception=False).

        Raises:
            exceptions.StatusBarAssertionError: Если проверка не пройдена (и raise_exception=True).
            exceptions.StatusBarException: Если произошла ошибка при чтении статус-бара (кроме "не найдено").
            ValueError: Если не задано ни одного критерия для проверки.
        """
        if expected_type is None and expected_id is None and expected_number is None and expected_text is None:
            raise ValueError("Необходимо задать хотя бы один критерий для проверки статус-бара.")

        start_time = time.monotonic()
        last_status_data: Optional[Tuple[str, str, str]] = None

        while time.monotonic() - start_time < timeout:
            try:
                current_status_data = self.get_status_message(window_id)
                if current_status_data:
                    last_status_data = current_status_data # Запоминаем последнее непустое сообщение
                    msg_type, msg_id, msg_number, msg_text = current_status_data

                    # --- Выполняем проверки ---
                    type_match = True
                    if expected_type is not None:
                        allowed_types = [expected_type] if isinstance(expected_type, str) else expected_type
                        type_match = msg_type in allowed_types
                    # --- НОВОЕ: Добавляем проверку ID ---
                    id_match = True
                    if expected_id is not None:
                        allowed_ids = [expected_id] if isinstance(expected_id, str) else expected_id
                        # Сравнение регистрозависимо, как и в SAP
                        id_match = msg_id in allowed_ids
                    # --- КОНЕЦ НОВОГО ---
                    number_match = True
                    if expected_number is not None:
                        allowed_numbers_str = []
                        if isinstance(expected_number, (str, int)):
                             allowed_numbers_str.append(str(expected_number))
                        else: # Список
                             allowed_numbers_str = [str(n) for n in expected_number]
                        # Сравниваем как строки, т.к. msg_number тоже строка
                        number_match = msg_number in allowed_numbers_str

                    text_match = True
                    if expected_text is not None:
                        cleaned_msg_text = msg_text.strip()
                        if isinstance(expected_text, re.Pattern):
                            text_match = expected_text.search(cleaned_msg_text) is not None
                        else: # Точное совпадение строки
                            text_match = cleaned_msg_text == str(expected_text).strip()

                    # --- Финальное решение ---
                    if type_match and id_match and number_match and text_match:
                        log.info(f"Проверка статус-бара пройдена. Сообщение: (T='{msg_type}', ID='{msg_id}', N='{msg_number}', Text='{msg_text}')")
                        return True # Все указанные критерии совпали

                # Если сообщение есть, но не совпало, или сообщения пока нет, ждем дальше
                time.sleep(poll_interval)

            except exceptions.StatusBarException as e:
                # Перебрасываем серьезные ошибки чтения
                if raise_exception: raise
                else: return False
            except Exception as e_loop:
                 # Другие возможные ошибки в цикле
                 msg = f"Неожиданная ошибка в цикле assert_status_bar: {e_loop}"
                 log.error(msg)
                 if raise_exception: raise exceptions.StatusBarAssertionError(msg) from e_loop
                 else: return False


        # --- Таймаут истек ---
        if last_status_data:
            # Сообщение было, но не подошло
            msg = (f"Проверка статус-бара НЕ пройдена. Последнее сообщение: {last_status_data}. "
                   f"Ожидалось: Type={expected_type}, ID={expected_id}, Number={expected_number}, Text='{expected_text}'.")
            log.warning(msg)
            if raise_exception: raise exceptions.StatusBarAssertionError(msg)
            return False
        else:
            # Сообщение так и не появилось
            if fail_on_timeout:
                msg = f"Проверка статус-бара НЕ пройдена. Сообщение не появилось в '{window_id}/sbar' за {timeout} сек."
                log.warning(msg)
                if raise_exception: raise exceptions.StatusBarAssertionError(msg)
                return False
            else:
                 # Не считать ошибкой, если сообщение не появилось (fail_on_timeout=False)
                 log.info(f"Сообщение в статус-баре не появилось за {timeout} сек (fail_on_timeout=False). Проверка пропущена.")
                 return True # Считаем успехом, т.к. нет сообщения для проверки            

    def _find_menu_item_recursive(self, menu_element: win32com.client.CDispatch, target_names: List[str]) -> Optional[win32com.client.CDispatch]:
        """Recursive helper to find a menu item by name."""
        try:
            element_text = getattr(menu_element, "Text", "")
            cleaned_text = element_text.replace("&", "")
            if cleaned_text in target_names:
                return menu_element
            if hasattr(menu_element, "Children"):
                 children_count = getattr(menu_element.Children, "Count", 0)
                 for i in range(children_count):
                    child = menu_element.Children(i)
                    found = self._find_menu_item_recursive(child, target_names)
                    if found:
                        return found
        except Exception:
             pass
        return None

    def select_menu_item_by_name(self, menu_path: List[str], window_id: str = "wnd[0]") -> None:
        """ Selects a menu item by navigating through menu names. """
        if not menu_path:
            raise ValueError("Menu path cannot be empty.")
        menu_bar_id = f"{window_id}/mbar"
        try:
            current_element = self.session_handle.findById(menu_bar_id)
            element_to_select = None
            for i, name in enumerate(menu_path):
                target_names = [name]
                found_element = self._find_menu_item_recursive(current_element, target_names)
                if not found_element:
                    raise exceptions.MenuNotFoundException(
                        f"Menu item '{name}' not found in path: {' -> '.join(menu_path[:i+1])}"
                    )
                if i == len(menu_path) - 1:
                    element_to_select = found_element
                else:
                     # Need to ensure the submenu is accessible for the next iteration.
                     # Selecting intermediate menus might be one way, but risky.
                     # Let's assume the recursive search handles expanded/unexpanded state.
                     current_element = found_element

            if element_to_select:
                 if hasattr(element_to_select, "Select"):
                     element_to_select.Select()
                 # --- ДОБАВЛЕНО: Попытка нажать кнопку, если Select нет (для некоторых меню) ---
                 elif hasattr(element_to_select, "Press"):
                      print(f"Warning: Menu item '{menu_path[-1]}' does not support Select(). Attempting Press().")
                      element_to_select.Press()
                 # --- КОНЕЦ ДОБАВЛЕНИЯ ---
                 else:
                      raise exceptions.ActionException(f"Found menu element for '{menu_path[-1]}' but it supports neither Select() nor Press(). Type: {getattr(element_to_select, 'Type', 'N/A')}")
            else:
                 raise exceptions.MenuNotFoundException(f"Could not resolve final menu item for path: {' -> '.join(menu_path)}")
        except exceptions.MenuNotFoundException:
            raise
        except Exception as e:
            raise exceptions.ActionException(f"Error selecting menu item by name {' -> '.join(menu_path)}: {e}")

    def start_transaction_robust(self, transaction: str, check_errors: bool = True) -> None:
        """ Starts a transaction using /N prefix and optionally checks status bar for common errors. """
        okcode_field = "wnd[0]/tbar[0]/okcd"
        tcode_command = f"/n{transaction}" # Use lowercase /n for safety
        try:
            self.write(okcode_field, tcode_command)
            self.press("wnd[0]/tbar[0]/btn[0]") # Use Enter button press
            if check_errors:
                sleep(0.5) # Wait for status bar update
                status = self.get_status_message()
                if status:
                    msg_type, msg_number, msg_text = status
                    if msg_type == 'E' or msg_type == 'A': # Error or Abort
                        # Common error codes (may vary slightly by system version)
                        if msg_number in ["00343", "343"]: # Transaction & does not exist
                            raise exceptions.TransactionNotFoundError(f"Transaction '{transaction}' not found. SAP: {msg_text}")
                        elif msg_number in ["00077", "077"]: # User & is not authorized...
                             raise exceptions.AuthorizationError(f"Not authorized for transaction '{transaction}'. SAP: {msg_text}")
                        elif msg_number in ["00410", "410"]: # Action was blocked... (e.g., SM04 lock)
                             raise exceptions.ActionBlockedError(f"Action blocked in transaction '{transaction}'. SAP: {msg_text}")
                        # --- ДОБАВЛЕНО: Еще одна проверка авторизации ---
                        elif msg_number == "00057": # No authorization to start transaction &
                              raise exceptions.AuthorizationError(f"Not authorized for transaction '{transaction}' (Msg 00057). SAP: {msg_text}")
                        # --- КОНЕЦ ДОБАВЛЕНИЯ ---
                        else: # General Error/Abort not specifically identified
                            print(f"Warning/Error in status bar after starting '{transaction}': Type={msg_type}, Num={msg_number}, Text={msg_text}")
                            # Optionally raise a generic error
                            # raise exceptions.TransactionException(f"Error after starting '{transaction}'. SAP: {msg_text}")
                    elif msg_type == 'S' and msg_number == "00344": # Transaction & is locked (SM01)
                         raise exceptions.ActionBlockedError(f"Transaction '{transaction}' is locked (SM01). SAP: {msg_text}")

        except (exceptions.TransactionNotFoundError, exceptions.AuthorizationError, exceptions.ActionBlockedError):
            raise
        except Exception as e:
            raise exceptions.ActionException(f"Error starting transaction '{transaction}': {e}")

    def iterate_elements_by_template(self,
                                     root_element_id: str,
                                     id_template: str,
                                     start_index: int,
                                     max_index: int = 50) -> Generator[Tuple[int, win32com.client.CDispatch], None, None]:
        """ Iterates through GUI elements based on an ID template within a root element. """
        # Root element finding is removed as per original code, findById is relative to session
        # try:
        #     root_element = self.session_handle.findById(root_element_id)
        # except Exception as e:
        #      raise exceptions.ElementNotFoundException(f"Root element '{root_element_id}' not found: {e}")

        for index in range(start_index, max_index + 1):
            try:
                element_id = id_template.format(index=index)
                element = self.session_handle.findById(element_id)
                yield index, element
            except Exception as e:
                # Check if it's likely an 'element not found' error
                if "findById" in str(e) or "control could not be found" in str(e).lower():
                    break # End of list/grid for this template
                else:
                    # Re-raise unexpected errors
                    raise exceptions.SapGuiComException(f"Error finding element with template '{id_template}' at index {index}: {e}")

    def print_all_elements(self, root_element_id: str = "wnd[0]") -> None:
        """ Prints the IDs and types of all direct child elements of a specified root element. """
        print(f"--- Elements inside '{root_element_id}' ---")
        try:
            root_element = self.session_handle.findById(root_element_id)
            if not hasattr(root_element, "Children"):
                print(f"Element '{root_element_id}' (Type: {getattr(root_element, 'Type', 'N/A')}) has no Children attribute.")
                return
            children_count = getattr(root_element.Children, "Count", 0)
            if children_count == 0:
                print("(No children found)")
                return
            for i in range(children_count):
                try:
                    child = root_element.Children(i)
                    child_id = getattr(child, "Id", f"<Error getting ID for index {i}>")
                    child_type = getattr(child, "Type", "N/A")
                    child_name = getattr(child, "Name", "") # Name (not ID) sometimes useful
                    print(f"  Index: {i}, ID: {child_id} (Type: {child_type}, Name: '{child_name}')")
                except Exception as e_child:
                    print(f"  Index: {i}, Error accessing child element: {e_child}")
        except Exception as e_root:
            if "findById" in str(e_root) or "control could not be found" in str(e_root).lower():
                 raise exceptions.ElementNotFoundException(f"Root element '{root_element_id}' not found: {e_root}")
            else:
                 raise exceptions.SapGuiComException(f"Error getting children for '{root_element_id}': {e_root}")
        finally:
                 print(f"--- End of elements for '{root_element_id}' ---")

    def scroll_element(self, element_id: str, position: int) -> None:
        """ Scrolls the vertical scrollbar of a given element to a specific position. """
        try:
            element = self.session_handle.findById(element_id)
            if hasattr(element, "verticalScrollbar"):
                scrollbar = element.verticalScrollbar
                if hasattr(scrollbar, "position"):
                    scrollbar.position = position
                else:
                    raise exceptions.ActionException(f"Scrollbar for element '{element_id}' does not have a 'position' attribute.")
            else:
                 # Fallback: Try setting firstVisibleRow for TableControl
                 if element.Type == "GuiTableControl" and hasattr(element, "firstVisibleRow"):
                     element.firstVisibleRow = position
                     print(f"Used firstVisibleRow for GuiTableControl {element_id}")
                 else:
                     raise exceptions.ActionException(f"Element '{element_id}' (Type: {getattr(element, 'Type', 'N/A')}) has no controllable 'verticalScrollbar' or applicable fallback.")
        except AttributeError as ae:
             raise exceptions.ActionException(f"Attribute error while scrolling element '{element_id}': {ae}. Check element type and properties.")
        except Exception as e:
            raise exceptions.ActionException(f"Error scrolling element '{element_id}' to position {position}: {e}")


    # --- НОВЫЙ КОД: Работа со свойствами и DumpState ---
    def get_element_property(self, element_id: str, property_name: str) -> Any:
        """
        Gets the value of a specified property for a given element.

        Args:
            element_id (str): The ID of the SAP GUI element.
            property_name (str): The name of the property to retrieve (case-sensitive).

        Returns:
            Any: The value of the property.

        Raises:
            exceptions.ElementNotFoundException: If the element cannot be found by ID.
            exceptions.PropertyNotFoundException: If the element does not have the specified property.
            exceptions.SapGuiComException: For other COM errors.
        """
        try:
            element = self.session_handle.findById(element_id)
            if hasattr(element, property_name):
                return getattr(element, property_name)
            else:
                # Check common case variations if initial getattr fails (experimental)
                if hasattr(element, property_name.lower()):
                    return getattr(element, property_name.lower())
                elif hasattr(element, property_name.capitalize()):
                     return getattr(element, property_name.capitalize())
                else:
                     raise exceptions.PropertyNotFoundException(
                         f"Property '{property_name}' not found for element '{element_id}' (Type: {element.Type}).")
        except exceptions.PropertyNotFoundException:
             raise
        except Exception as e:
            if "findById" in str(e) or "control could not be found" in str(e).lower():
                raise exceptions.ElementNotFoundException(f"Element '{element_id}' not found: {e}")
            else:
                raise exceptions.SapGuiComException(f"Error getting property '{property_name}' for element '{element_id}': {e}")

    def set_element_property(self, element_id: str, property_name: str, value: Any) -> None:
        """
        Sets the value of a specified property for a given element.
        Warning: Use with caution. Not all properties are writable or intended to be changed.

        Args:
            element_id (str): The ID of the SAP GUI element.
            property_name (str): The name of the property to set (case-sensitive).
            value (Any): The value to assign to the property.

        Raises:
            exceptions.ElementNotFoundException: If the element cannot be found by ID.
            exceptions.PropertyNotFoundException: If the element does not have the specified property.
            exceptions.SapGuiComException: If setting the property fails (e.g., read-only) or other COM errors.
        """
        try:
            element = self.session_handle.findById(element_id)
            if not hasattr(element, property_name):
                 raise exceptions.PropertyNotFoundException(
                     f"Property '{property_name}' not found for element '{element_id}' (Type: {element.Type}). Cannot set value.")
            setattr(element, property_name, value)
        except exceptions.PropertyNotFoundException:
             raise
        except Exception as e:
            if "findById" in str(e) or "control could not be found" in str(e).lower():
                raise exceptions.ElementNotFoundException(f"Element '{element_id}' not found: {e}")
            else:
                # Error might indicate property is read-only
                raise exceptions.SapGuiComException(f"Error setting property '{property_name}' for element '{element_id}': {e}")

    def _dump_recursive(self, com_object: win32com.client.CDispatch, current_depth: int, max_depth: int) -> Dict[str, Any]:
        """Internal recursive helper for dump_element_state."""
        if current_depth > max_depth:
            return {"..." : f"Max depth ({max_depth}) reached"}

        state = {}
        # Basic properties that are usually safe to read
        safe_props = ['Id', 'Type', 'Name', 'Text', 'Changeable', 'ContainerType', 'ScreenLeft', 'ScreenTop', 'Width', 'Height', 'Tooltip', 'DefaultTooltip', 'IconName']
        for prop in safe_props:
            try:
                if hasattr(com_object, prop):
                    state[prop] = getattr(com_object, prop)
                # else:
                #     state[prop] = "<Not Available>"
            except Exception as e:
                state[prop] = f"<Error Reading: {e}>"

        # Attempt to read other properties (use with caution) - limited list for safety
        other_props = ['selected', 'Left', 'Top'] # Add more cautiously if needed
        for prop in other_props:
             try:
                 if hasattr(com_object, prop):
                      state[prop] = getattr(com_object, prop)
             except Exception as e:
                  state[prop] = f"<Error Reading: {e}>"


        # Handle Children recursively
        if hasattr(com_object, "Children"):
            try:
                children_count = getattr(com_object.Children, "Count", 0)
                if children_count > 0:
                    state["Children"] = []
                    # Limit number of children dumped to avoid excessive output
                    max_children_to_dump = 10
                    for i in range(min(children_count, max_children_to_dump)):
                         try:
                             child = com_object.Children(i)
                             child_state = self._dump_recursive(child, current_depth + 1, max_depth)
                             state["Children"].append(child_state)
                         except Exception as e_child:
                              state["Children"].append({"Index": i, "Error": f"<Error Accessing Child: {e_child}>"})
                    if children_count > max_children_to_dump:
                         state["Children"].append({"..." : f"{children_count - max_children_to_dump} more children not shown"})

            except Exception as e_children:
                state["Children"] = f"<Error Accessing Children: {e_children}>"

        return state

    def dump_element_state(self, element_id: str, recursive: bool = True, max_depth: int = 3, print_output: bool = True) -> Optional[Dict[str, Any]]:
        """
        Retrieves and optionally prints the state (properties and children) of an element.
        Useful for debugging and understanding element structure. Now uses _build_snapshot_recursive.

        Args:
            element_id (str): The ID of the SAP GUI element to dump.
            recursive (bool): If True, recursively dumps child elements up to max_depth.
            max_depth (int): The maximum depth for recursive dumping.
            print_output (bool): If True, pretty-prints the dumped state to the console.

        Returns:
            Optional[Dict[str, Any]]: A dictionary representing the element's state,
                                      or None if print_output is True.

        Raises:
            exceptions.ElementNotFoundException: If the root element cannot be found.
            exceptions.SapGuiComException: For other COM errors during dumping.
        """
        try:
            root_element = self.session_handle.findById(element_id)
            actual_max_depth = max_depth if recursive else 0

            # Используем новый рекурсивный построитель, но с ограниченным набором свойств по умолчанию для скорости
            default_props_for_dump = ['Id', 'Type', 'Name', 'Text', 'Tooltip', 'Changeable', 'ClassName']
            dump_data = self._build_snapshot_recursive(
                com_object=root_element,
                current_depth=0,
                max_depth=actual_max_depth,
                props_include=default_props_for_dump, # Ограничиваем свойства для dump_element_state
                props_exclude=[]
            )

            if print_output:
                print(f"--- Dump State for Element '{element_id}' (Max Depth: {actual_max_depth}) ---")
                pprint.pprint(dump_data, indent=2)
                print(f"--- End Dump State for Element '{element_id}' ---")
                return None
            else:
                return dump_data

        except exceptions.ElementNotFoundException:
            raise # Перебрасываем как есть
        except Exception as e:
            # Ловим другие ошибки при вызове findById или _build_snapshot_recursive
            raise exceptions.SapGuiComException(f"Error dumping state for element '{element_id}': {e}") from e


    # --- НОВЫЙ КОД: Метод для получения GuiTree ---
    def get_tree(self, element_id: str) -> GuiTree:
        """
        Gets a GuiTree object for interacting with a SAP GUI tree element.

        Args:
            element_id (str): The ID of the GuiTree element (e.g., GuiShell of subtype Tree).

        Returns:
            GuiTree: An object for tree operations.

        Raises:
            exceptions.ElementNotFoundException: If the element is not found.
            exceptions.InvalidElementTypeException: If the element is not a GuiTree compatible type.
            exceptions.SapGuiComException: For other COM errors.

        Example:
            ```python
            try:
                tree = main_window.get_tree("wnd[0]/usr/...")
                all_keys = tree.get_all_node_keys()
                tree.select_node(all_keys[0])
            except exceptions.ActionException as e:
                print(f"Error interacting with tree: {e}")
            ```
        """
        # GuiTree.__init__ handles findById and type checking
        return GuiTree(self.session_handle, element_id)
    # --- КОНЕЦ НОВОГО КОДА ---
    
    def handle_unexpected_popup(self,
                                popup_ids: Optional[List[str]] = None,
                                # --- Изменения для кнопки "Нет" ---
                                press_no_button_id: Optional[str] = None, # ID кнопки "Нет" (напр., "usr/btnSPOP-OPTION2")
                                # --- Остальные параметры ---
                                press_button_id: Optional[str] = None, # ID основной кнопки (напр., "tbar[0]/btn[0]", "usr/btnSPOP-OPTION1")
                                action_vkey: Optional[int] = 0, # VKey для отправки (0=Enter), если кнопки не найдены/не указаны
                                wait_after_action: float = 0.5,
                                log_details: bool = True) -> bool:
        """
        Обнаруживает и ПЫТАЕТСЯ обработать простые неожиданные всплывающие окна.
        Проверяет указанные ID окон. Если найдено, пытается нажать указанную кнопку или отправить VKey.
        Приоритет действий: press_no_button_id -> press_button_id -> action_vkey.

        Args:
            popup_ids (Optional[List[str]]): Список ID окон для проверки (напр., ["wnd[1]", "wnd[2]"]).
                                             По умолчанию: ["wnd[1]", "wnd[2]"].
            press_no_button_id (Optional[str]): ID кнопки "Нет" (или аналогичной отменяющей) для нажатия.
                                                Имеет НАИВЫСШИЙ приоритет. Стандартно: "usr/btnSPOP-OPTION2".
            press_button_id (Optional[str]): ID основной (обычно подтверждающей) кнопки для нажатия.
                                             Например: "tbar[0]/btn[0]" (ОК), "usr/btnSPOP-OPTION1" (Да).
                                             Используется, если press_no_button_id не указан или не найден.
            action_vkey (Optional[int]): VKey для отправки в найденное окно, если ни одна из указанных кнопок не найдена.
                                         По умолчанию: 0 (Enter). Установите None, чтобы не отправлять VKey как fallback.
            wait_after_action (float): Пауза в секундах ПОСЛЕ успешного действия.
            log_details (bool): Логировать ли обнаружение и действия.

        Returns:
            bool: True, если всплывающее окно было обнаружено И успешно обработано.
                  False в противном случае.
        """
        popup_ids = popup_ids if popup_ids is not None else ["wnd[1]", "wnd[2]"]
        handled_successfully = False

        for popup_id in popup_ids:
            try:
                element_exists = False
                try:
                    self.session_handle.FindById(popup_id, False)
                    element_exists = True
                except Exception:
                    element_exists = False

                if element_exists:
                    if log_details:
                        popup_title = ""
                        try: popup_title = self.get_element_property(popup_id, "Text")
                        except: pass
                        log.warning(f"Обнаружено возможное всплывающее окно: {popup_id} (Заголовок: '{popup_title}')")

                    action_taken = False

                    # 1. Приоритет: Попытка нажать кнопку "Нет" (если указана)
                    if press_no_button_id:
                        full_no_button_id = f"{popup_id}/{press_no_button_id}"
                        try:
                            button_exists = False
                            try:
                                self.session_handle.FindById(full_no_button_id, False)
                                button_exists = True
                            except: pass

                            if button_exists:
                                self.press(full_no_button_id)
                                if log_details:
                                    log.info(f"Нажата кнопка 'Нет' (или аналог) '{full_no_button_id}' во всплывающем окне {popup_id}.")
                                action_taken = True # Устанавливаем флаг, что действие выполнено
                            elif log_details:
                                log.debug(f"Кнопка 'Нет' '{full_no_button_id}' не найдена в окне {popup_id}.")
                        except Exception as btn_e:
                            if log_details:
                                log.error(f"Ошибка при попытке нажать кнопку 'Нет' '{full_no_button_id}': {btn_e}")

                    # 2. Попытка нажать основную кнопку (если "Нет" не нажималась и кнопка указана)
                    if not action_taken and press_button_id:
                        full_button_id = f"{popup_id}/{press_button_id}"
                        try:
                            button_exists = False
                            try:
                                self.session_handle.FindById(full_button_id, False)
                                button_exists = True
                            except: pass

                            if button_exists:
                                self.press(full_button_id)
                                if log_details:
                                    log.info(f"Нажата основная кнопка '{full_button_id}' во всплывающем окне {popup_id}.")
                                action_taken = True
                            elif log_details:
                                log.debug(f"Основная кнопка '{full_button_id}' не найдена в окне {popup_id}.")
                        except Exception as btn_e:
                            if log_details:
                                log.error(f"Ошибка при попытке нажать основную кнопку '{full_button_id}': {btn_e}")

                    # 3. Попытка отправить VKey (если никакая кнопка не нажималась и VKey указан)
                    if not action_taken and action_vkey is not None:
                        try:
                            self.send_v_key(element=popup_id, value=action_vkey)
                            if log_details:
                                log.info(f"Отправлен VKey {action_vkey} во всплывающее окно {popup_id}.")
                            action_taken = True
                        except Exception as vkey_e:
                            if log_details:
                                log.error(f"Ошибка при отправке VKey {action_vkey} в окно {popup_id}: {vkey_e}")

                    # Финальная проверка и выход из цикла, если успешно обработали
                    if action_taken:
                        handled_successfully = True
                        if wait_after_action > 0:
                            if log_details:
                                log.debug(f"Пауза {wait_after_action} сек после обработки окна {popup_id}.")
                            time.sleep(wait_after_action)
                        break # Обработали первое найденное окно, выходим
                    elif log_details:
                         # Это сообщение теперь менее критично, т.к. могло быть не найдено указанных кнопок
                        log.debug(f"Не удалось выполнить настроенное действие (кнопки/VKey) для окна {popup_id}.")

            except Exception as outer_e:
                if log_details:
                    log.error(f"Критическая ошибка при проверке/обработке окна {popup_id}: {outer_e}")

        return handled_successfully

    def save_gui_snapshot(self,
                          filepath: Union[str, Path],
                          root_element_id: str = "wnd[0]",
                          max_depth: Optional[int] = None, # None = Без ограничений
                          properties_to_include: Optional[List[str]] = None, # None = Все доступные
                          properties_to_exclude: Optional[List[str]] = None, # Свойства, которые не надо сохранять
                          output_format: str = 'json',
                          include_children: bool = True
                          ) -> None:
        """
        Создает "слепок" (snapshot) иерархии и свойств GUI-элементов,
        начиная с указанного корневого элемента, и сохраняет его в файл.

        Args:
            filepath (Union[str, Path]): Путь к файлу для сохранения (JSON или YAML).
            root_element_id (str): ID корневого элемента для начала сканирования. По умолчанию "wnd[0]".
            max_depth (Optional[int]): Максимальная глубина рекурсии. None - без ограничений.
            properties_to_include (Optional[List[str]]): Список имен свойств для включения.
                                                         Если None, пытается включить все найденные свойства
                                                         (кроме исключенных).
            properties_to_exclude (Optional[List[str]]): Список имен свойств для явного исключения.
                                                         Полезно для пропуска "шумных" или проблемных свойств.
            output_format (str): Формат вывода: 'json' или 'yaml'. По умолчанию 'json'.
            include_children (bool): Включать ли дочерние элементы рекурсивно. По умолчанию True.

        Raises:
            exceptions.ElementNotFoundException: Если корневой элемент не найден.
            exceptions.SapGuiComException: При других ошибках COM во время сканирования.
            ValueError: Если указан неверный output_format или yaml не установлен.
            IOError: При ошибках записи файла.
        """
        log.info(f"Создание слепка GUI для элемента '{root_element_id}' -> {filepath}")
        start_time = time.time()

        if output_format.lower() not in ['json', 'yaml']:
            raise ValueError("Неверный output_format. Допустимые значения: 'json', 'yaml'.")
        if output_format.lower() == 'yaml' and yaml is None:
            raise ValueError("Для формата 'yaml' необходимо установить библиотеку PyYAML: pip install pyyaml")

        # --- Получаем корневой COM-объект ---
        try:
            root_com_object = self.session_handle.findById(root_element_id)
        except Exception as e:
            if "findById" in str(e) or "control could not be found" in str(e).lower():
                raise exceptions.ElementNotFoundException(f"Корневой элемент '{root_element_id}' не найден: {e}")
            else:
                raise exceptions.SapGuiComException(f"Ошибка доступа к корневому элементу '{root_element_id}': {e}")

        # --- Запускаем рекурсивный сбор данных ---
        try:
            snapshot_data = self._build_snapshot_recursive(
                com_object=root_com_object,
                current_depth=0,
                max_depth=max_depth if include_children else 0, # Если дети не нужны, глубина 0
                props_include=properties_to_include,
                props_exclude=properties_to_exclude if properties_to_exclude else [] # Убедимся, что это список
            )
        except Exception as e_dump:
             # Ловим ошибки, возникшие при рекурсивном обходе
             log.exception(f"Ошибка во время рекурсивного сбора данных для слепка: {e_dump}")
             raise exceptions.SapGuiComException(f"Ошибка сбора данных для слепка: {e_dump}") from e_dump

        # --- Сохраняем в файл ---
        file_path_obj = Path(filepath)
        try:
            # Создаем родительские директории
            file_path_obj.parent.mkdir(parents=True, exist_ok=True)

            with open(file_path_obj, 'w', encoding='utf-8') as f:
                if output_format.lower() == 'json':
                    json.dump(snapshot_data, f, indent=2, ensure_ascii=False, default=str) # default=str для несериализуемых типов
                elif output_format.lower() == 'yaml':
                    # Используем Dumper=yaml.SafeDumper или просто dump, если безопасность не критична
                    # allow_unicode=True важно для не-ASCII символов
                    yaml.dump(snapshot_data, f, allow_unicode=True, indent=2, default_flow_style=False, sort_keys=False)

            end_time = time.time()
            log.info(f"Слепок GUI успешно сохранен в '{file_path_obj.resolve()}'. Время: {end_time - start_time:.2f} сек.")

        except IOError as e_io:
            log.error(f"Ошибка записи файла слепка '{filepath}': {e_io}")
            raise
        except Exception as e_save:
            log.error(f"Неожиданная ошибка при сохранении файла слепка '{filepath}': {e_save}")
            raise
        
    def _build_snapshot_recursive(self,
                                  com_object: Any,
                                  current_depth: int,
                                  max_depth: Optional[int],
                                  props_include: Optional[List[str]],
                                  props_exclude: List[str]
                                  ) -> Dict[str, Any]:
        """Внутренний рекурсивный метод для построения словаря слепка."""
        if max_depth is not None and current_depth > max_depth:
            # Если есть ID, возвращаем его и маркер глубины
            try: id_val = getattr(com_object, "Id", "<No ID>")
            except: id_val = "<Error Reading ID>"
            return {"Id": id_val, "__Status__": f"Max depth ({max_depth}) reached"}

        element_data: Dict[str, Any] = {}
        property_names_to_try: List[str] = []

        # --- Определяем, какие свойства читать ---
        if props_include is not None:
            property_names_to_try = [p for p in props_include if p not in props_exclude]
        else:
            # Пытаемся получить все атрибуты через dir(), фильтруем ненужные
            try:
                all_attrs = dir(com_object)
                potential_props = [
                    p for p in all_attrs
                    if not p.startswith('_') # Убираем внутренние
                    and p not in props_exclude # Убираем явно исключенные
                    # Попытка отфильтровать методы (не всегда надежно для COM)
                    # and not callable(getattr(com_object, p, None))
                ]
                # Добавляем стандартные важные свойства, если их нет в dir()
                standard_props = ['Id', 'Type', 'Name', 'Text', 'Tooltip', 'Changeable', 'ContainerType', 'ScreenLeft', 'ScreenTop', 'Width', 'Height', 'DefaultTooltip', 'IconName', 'ClassName']
                for sp in standard_props:
                     if sp not in potential_props and sp not in props_exclude:
                         potential_props.append(sp)
                property_names_to_try = sorted(list(set(potential_props))) # Уникальные и сортированные

            except Exception as e_dir:
                log.warning(f"Не удалось получить атрибуты через dir() для объекта: {e_dir}. Используем только стандартный набор.")
                # Fallback на стандартный набор, если dir() не сработал
                standard_props = ['Id', 'Type', 'Name', 'Text', 'Tooltip', 'Changeable', 'ContainerType', 'ScreenLeft', 'ScreenTop', 'Width', 'Height', 'DefaultTooltip', 'IconName', 'ClassName']
                property_names_to_try = [p for p in standard_props if p not in props_exclude]


        # --- Читаем выбранные свойства ---
        for prop_name in property_names_to_try:
            try:
                # Используем hasattr для проверки перед чтением (доп. защита)
                if hasattr(com_object, prop_name):
                    value = getattr(com_object, prop_name)
                    # Пытаемся обработать COM коллекции (простой случай)
                    if isinstance(value, win32com.client.CDispatch) and hasattr(value, 'Count') and hasattr(value, 'Item'):
                         try:
                              element_data[prop_name] = [value.Item(i) for i in range(value.Count)]
                         except Exception:
                              element_data[prop_name] = "<Error Reading COM Collection>"
                    else:
                         element_data[prop_name] = value
                # else: # Не добавляем свойство, если его нет
                #    pass
            except Exception as e_prop:
                element_data[prop_name] = f"<Error Reading Property: {type(e_prop).__name__}>"

        # --- Рекурсивно обрабатываем дочерние элементы ---
        if (max_depth is None or current_depth < max_depth) and hasattr(com_object, "Children"):
            try:
                children_collection = getattr(com_object, "Children", None)
                if children_collection and hasattr(children_collection, "Count"):
                    children_count = children_collection.Count
                    if children_count > 0:
                        element_data["Children"] = []
                        for i in range(children_count):
                            try:
                                child_com_object = children_collection(i) # Доступ к дочернему элементу
                                child_data = self._build_snapshot_recursive(
                                    com_object=child_com_object,
                                    current_depth=current_depth + 1,
                                    max_depth=max_depth,
                                    props_include=props_include,
                                    props_exclude=props_exclude
                                )
                                element_data["Children"].append(child_data)
                            except Exception as e_child:
                                element_data["Children"].append({"__Index__": i, "__Error__": f"<Error Accessing/Processing Child: {type(e_child).__name__}>"})
            except Exception as e_children:
                element_data["Children"] = f"<Error Accessing Children Collection: {type(e_children).__name__}>"

        return element_data
    def save_gui_snapshot_from_schema(self,
                                      filepath: Union[str, Path],
                                      object_schema: Dict[str, Any], # Загруженный sap_gui_objects.json
                                      # enum_schema: Dict[str, Any], # Пока не используется напрямую для свойств
                                      root_element_id: str = "wnd[0]",
                                      max_depth: Optional[int] = None, # None = Без ограничений
                                      properties_to_exclude: Optional[List[str]] = None, # Свойства, которые не надо сохранять
                                      output_format: str = 'json',
                                      include_children: bool = True
                                      ) -> None:
        """
        Создает "слепок" GUI-элементов, используя ЗАДАННУЮ СХЕМУ для определения
        свойств, которые нужно попытаться прочитать.

        Args:
            filepath (Union[str, Path]): Путь к файлу для сохранения (JSON или YAML).
            object_schema (Dict[str, Any]): Словарь, загруженный из JSON-файла с описанием
                                            методов и свойств объектов (sap_gui_objects.json).
            root_element_id (str): ID корневого элемента для начала сканирования. По умолчанию "wnd[0]".
            max_depth (Optional[int]): Максимальная глубина рекурсии. None - без ограничений.
            properties_to_exclude (Optional[List[str]]): Список имен свойств для явного исключения
                                                         из чтения (даже если они есть в схеме).
            output_format (str): Формат вывода: 'json' или 'yaml'. По умолчанию 'json'.
            include_children (bool): Включать ли дочерние элементы рекурсивно. По умолчанию True.

        Raises:
            exceptions.ElementNotFoundException: Если корневой элемент не найден.
            exceptions.SapGuiComException: При других ошибках COM во время сканирования.
            ValueError: Если указан неверный output_format, yaml не установлен, или схема не передана.
            KeyError: Если тип элемента не найден в переданной object_schema.
            IOError: При ошибках записи файла.
        """
        log.info(f"Создание слепка GUI по СХЕМЕ для элемента '{root_element_id}' -> {filepath}")
        start_time = time.time()

        if not object_schema:
             raise ValueError("Необходимо передать object_schema (загруженный sap_gui_objects.json).")
        if output_format.lower() not in ['json', 'yaml']:
            raise ValueError("Неверный output_format. Допустимые значения: 'json', 'yaml'.")
        if output_format.lower() == 'yaml' and yaml is None:
            raise ValueError("Для формата 'yaml' необходимо установить библиотеку PyYAML: pip install pyyaml")

        props_exclude_list = properties_to_exclude if properties_to_exclude else []

        # --- Получаем корневой COM-объект ---
        try:
            root_com_object = self.session_handle.findById(root_element_id)
        except Exception as e:
            if "findById" in str(e) or "control could not be found" in str(e).lower():
                raise exceptions.ElementNotFoundException(f"Корневой элемент '{root_element_id}' не найден: {e}")
            else:
                raise exceptions.SapGuiComException(f"Ошибка доступа к корневому элементу '{root_element_id}': {e}")

        # --- Запускаем рекурсивный сбор данных по схеме ---
        try:
            snapshot_data = self._build_snapshot_from_schema_recursive(
                com_object=root_com_object,
                current_depth=0,
                max_depth=max_depth if include_children else 0,
                object_schema=object_schema,
                props_exclude=props_exclude_list
            )
        except Exception as e_dump:
             log.exception(f"Ошибка во время рекурсивного сбора данных для слепка по схеме: {e_dump}")
             raise exceptions.SapGuiComException(f"Ошибка сбора данных для слепка по схеме: {e_dump}") from e_dump

        # --- Сохраняем в файл ---
        file_path_obj = Path(filepath)
        try:
            file_path_obj.parent.mkdir(parents=True, exist_ok=True)
            with open(file_path_obj, 'w', encoding='utf-8') as f:
                if output_format.lower() == 'json':
                    json.dump(snapshot_data, f, indent=2, ensure_ascii=False, default=str)
                elif output_format.lower() == 'yaml':
                    yaml.dump(snapshot_data, f, allow_unicode=True, indent=2, default_flow_style=False, sort_keys=False)

            end_time = time.time()
            log.info(f"Слепок GUI по СХЕМЕ успешно сохранен в '{file_path_obj.resolve()}'. Время: {end_time - start_time:.2f} сек.")

        except IOError as e_io:
            log.error(f"Ошибка записи файла слепка '{filepath}': {e_io}")
            raise
        except Exception as e_save:
            log.error(f"Неожиданная ошибка при сохранении файла слепка '{filepath}': {e_save}")
            raise


    def _build_snapshot_from_schema_recursive(self,
                                              com_object: Any,
                                              current_depth: int,
                                              max_depth: Optional[int],
                                              object_schema: Dict[str, Any],
                                              props_exclude: List[str]
                                              ) -> Dict[str, Any]:
        """Внутренний рекурсивный метод для построения словаря слепка ПО СХЕМЕ."""
        if max_depth is not None and current_depth > max_depth:
            try: id_val = getattr(com_object, "Id", "<No ID>")
            except: id_val = "<Error Reading ID>"
            return {"Id": id_val, "__Status__": f"Max depth ({max_depth}) reached"}

        element_data: Dict[str, Any] = {}
        element_type: Optional[str] = None
        schema_properties: List[Dict[str, Any]] = [] # Список словарей свойств из схемы

        # --- Пытаемся определить тип элемента ---
        try:
            element_type = getattr(com_object, "Type")
            element_data["Type"] = element_type # Сохраняем тип в любом случае
        except Exception as e_type:
            element_data["Type"] = f"<Error Reading Type: {type(e_type).__name__}>"
            log.warning(f"Не удалось прочитать тип элемента: {e_type}")
            # Не можем продолжить поиск по схеме без типа

        # --- Ищем свойства в схеме по типу ---
        if element_type:
            schema_key = f"{element_type} Object" # Формируем ключ для поиска в схеме
            type_schema = object_schema.get(schema_key)

            if type_schema and isinstance(type_schema.get("properties"), list):
                schema_properties = type_schema["properties"]
                log.debug(f"Найдены свойства в схеме для типа '{element_type}'.")
            else:
                log.warning(f"Схема для типа '{element_type}' (ключ '{schema_key}') не найдена или не содержит список 'properties' в object_schema.")
                # Fallback: Попытаемся прочитать хотя бы базовые свойства
                basic_props = ['Id', 'Name', 'Text', 'Tooltip', 'Changeable', 'ScreenLeft', 'ScreenTop', 'Width', 'Height']
                schema_properties = [{"name": p} for p in basic_props] # Формируем структуру как в схеме

        # --- Читаем свойства, определенные в схеме (или базовые) ---
        property_names_to_try = [
            prop.get("name") for prop in schema_properties
            if prop.get("name") and prop.get("name") not in props_exclude
        ]

        # Всегда пытаемся прочитать ID, даже если его нет в списке (он ключевой)
        if "Id" not in property_names_to_try and "Id" not in props_exclude:
             property_names_to_try.insert(0, "Id")

        for prop_name in property_names_to_try:
            try:
                # Используем get_element_property, который уже имеет защиту hasattr
                # или просто getattr с try-except
                # if hasattr(com_object, prop_name): # Доп. проверка не помешает
                value = getattr(com_object, prop_name)

                # Обработка COM коллекций (как в предыдущем методе)
                if isinstance(value, win32com.client.CDispatch) and hasattr(value, 'Count') and hasattr(value, 'Item'):
                    try:
                        # Ограничим количество элементов коллекции для производительности
                        count = value.Count
                        max_coll_items = 50
                        items = []
                        for i in range(min(count, max_coll_items)):
                             # Пытаемся получить простое представление элемента коллекции
                             item_val = value.Item(i)
                             if isinstance(item_val, win32com.client.CDispatch):
                                  # Если это сложный объект, берем его ID или текст
                                  items.append(getattr(item_val, 'Id', getattr(item_val, 'Text', str(item_val))))
                             else:
                                  items.append(item_val)

                        element_data[prop_name] = items
                        if count > max_coll_items:
                            element_data[prop_name].append(f"... ({count - max_coll_items} more items)")

                    except Exception as e_coll:
                         element_data[prop_name] = f"<Error Reading COM Collection '{prop_name}': {type(e_coll).__name__}>"
                else:
                    element_data[prop_name] = value
                # else:
                #    log.debug(f"Свойство '{prop_name}' отсутствует у объекта типа '{element_type}', хотя ожидалось по схеме.")

            except Exception as e_prop:
                element_data[prop_name] = f"<Error Reading Property: {type(e_prop).__name__}>"

        # --- Рекурсивно обрабатываем дочерние элементы ---
        if (max_depth is None or current_depth < max_depth) and hasattr(com_object, "Children"):
            try:
                children_collection = getattr(com_object, "Children", None)
                if children_collection and hasattr(children_collection, "Count"):
                    children_count = children_collection.Count
                    if children_count > 0:
                        element_data["Children"] = []
                        for i in range(children_count):
                            try:
                                child_com_object = children_collection(i)
                                child_data = self._build_snapshot_from_schema_recursive(
                                    com_object=child_com_object,
                                    current_depth=current_depth + 1,
                                    max_depth=max_depth,
                                    object_schema=object_schema,
                                    props_exclude=props_exclude
                                )
                                element_data["Children"].append(child_data)
                            except Exception as e_child:
                                element_data["Children"].append({"__Index__": i, "__Error__": f"<Error Accessing/Processing Child: {type(e_child).__name__}>"})
            except Exception as e_children:
                element_data["Children"] = f"<Error Accessing Children Collection: {type(e_children).__name__}>"

        return element_data

    def find_element_id_by_locator(self, locator_str: str, target_element_types: Optional[List[str]] = None) -> Optional[str]:
        """
        Находит ID элемента, используя семантический локатор.
        Это основной метод для разрешения локаторов перед вызовом действия.

        Args:
            locator_str: Строка локатора (e.g., "Пользователь", "@ Пароль", "=Сохранить").
            target_element_types: Опциональный список типов элементов для поиска.

        Returns:
            ID найденного элемента или None, если не найден.
        """
        if not self._finder:
            raise RuntimeError("Semantic locator support is not enabled. Pass a SapElementFinder instance when creating Window.")
        try:
            return self._finder.find_element(locator_str, target_element_types)
        except exceptions.SapGuiComException as e:
            log.error(f"Error during element finding process: {e}")
            return None

    def press_by_locator(self, locator_str: str, target_element_types: Optional[List[str]] = ["GuiButton", "GuiTab"]) -> None:
        """ Finds element by *semantic locator* and presses it. """
        element_id = self.find_element_id_by_locator(locator_str, target_element_types)
        if not element_id:
            raise exceptions.ElementNotFoundException(f"Element not found using locator: '{locator_str}' (Target types: {target_element_types})")
        log.info(f"Pressing element found by locator '{locator_str}': {element_id}")
        self.press(element_id) # Вызов оригинального метода с найденным ID

    def write_by_locator(self, locator_str: str, text: str, target_element_types: Optional[List[str]] = ["GuiTextField", "GuiCTextField", "GuiPasswordField", "GuiComboBox"]) -> None:
        """ Finds element by *semantic locator* and writes text into it. """
        element_id = self.find_element_id_by_locator(locator_str, target_element_types)
        if not element_id:
            raise exceptions.ElementNotFoundException(f"Element not found using locator: '{locator_str}' (Target types: {target_element_types})")
        log.info(f"Writing to element found by locator '{locator_str}': {element_id}")
        self.write(element_id, text) # Вызов оригинального метода

    def read_by_locator(self, locator_str: str, target_element_types: Optional[List[str]] = ["GuiTextField", "GuiCTextField", "GuiPasswordField", "GuiComboBox", "GuiLabel"]) -> str:
        """ Finds element by *semantic locator* and reads its text property. """
        element_id = self.find_element_id_by_locator(locator_str, target_element_types)
        if not element_id:
            raise exceptions.ElementNotFoundException(f"Element not found using locator: '{locator_str}' (Target types: {target_element_types})")
        log.info(f"Reading from element found by locator '{locator_str}': {element_id}")
        return self.read(element_id) # Вызов оригинального метода

    def select_by_locator(self, locator_str: str, target_element_types: Optional[List[str]] = ["GuiCheckBox", "GuiRadioButton", "GuiTab", "GuiMenu"]) -> None:
        """ Finds element by *semantic locator* and selects it. """
        element_id = self.find_element_id_by_locator(locator_str, target_element_types)
        if not element_id:
            raise exceptions.ElementNotFoundException(f"Element not found using locator: '{locator_str}' (Target types: {target_element_types})")
        log.info(f"Selecting element found by locator '{locator_str}': {element_id}")
        self.select(element_id) # Вызов оригинального метода

    def is_selected_by_locator(self, locator_str: str, target_element_types: Optional[List[str]] = ["GuiCheckBox", "GuiRadioButton"]) -> bool:
        """ Finds element by *semantic locator* and gets its selection status. """
        element_id = self.find_element_id_by_locator(locator_str, target_element_types)
        if not element_id:
            raise exceptions.ElementNotFoundException(f"Element not found using locator: '{locator_str}' (Target types: {target_element_types})")
        log.info(f"Getting selection status for element found by locator '{locator_str}': {element_id}")
        return self.is_selected(element_id) # Вызов оригинального метода

    def set_checkbox_by_locator(self, locator_str: str, selected: bool, target_element_types: Optional[List[str]] = ["GuiCheckBox"]) -> None:
        """ Finds checkbox by *semantic locator* and sets its status. """
        element_id = self.find_element_id_by_locator(locator_str, target_element_types)
        if not element_id:
            raise exceptions.ElementNotFoundException(f"Checkbox not found using locator: '{locator_str}' (Target types: {target_element_types})")
        log.info(f"Setting checkbox status for element found by locator '{locator_str}': {element_id}")
        self.set_checkbox(element_id, selected) # Вызов оригинального метода

    def visualize_by_locator(self, locator_str: str, seconds: int = 1, target_element_types: Optional[List[str]] = None) -> None:
        """ Finds element by *semantic locator* and visualizes it. """
        # Используем типы по умолчанию, если не заданы
        effective_types = target_element_types if target_element_types is not None else DEFAULT_TARGET_TYPES
        element_id = self.find_element_id_by_locator(locator_str, effective_types)
        if not element_id:
            raise exceptions.ElementNotFoundException(f"Element not found using locator: '{locator_str}' (Target types: {effective_types})")
        log.info(f"Visualizing element found by locator '{locator_str}': {element_id}")
        self.visualize(element_id, seconds) # Вызов оригинального метода

    def exists_by_locator(self, locator_str: str, target_element_types: Optional[List[str]] = None) -> bool:
        """ Checks if an element exists using a *semantic locator*. """
        element_id = self.find_element_id_by_locator(locator_str, target_element_types)
        return element_id is not None

    # --- Вспомогательный метод для обработки исключений (можно сделать приватным) ---
    def _handle_find_or_action_exception(self, action_description: str, element_desc: str, exception_obj: Exception):
        """ Преобразует COM ошибки в специфичные исключения sapscriptwizard. """
        error_str = str(exception_obj).lower()
        if "findById" in error_str or "control could not be found" in error_str or "element not found" in error_str:
             # Обертываем исходное исключение для сохранения трассировки
            raise exceptions.ElementNotFoundException(f"Element '{element_desc}' not found while {action_description}.") from exception_obj
        elif isinstance(exception_obj, exceptions.InvalidElementTypeException):
             raise # Перебрасываем как есть
        else:
             # Общая ошибка действия
            raise exceptions.ActionException(f"Error {action_description} element '{element_desc}': {exception_obj}") from exception_obj
    # --- NEW CODE END ---