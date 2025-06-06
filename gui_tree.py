# Файл: sapscriptwizard/gui_tree.py
import win32com.client
from typing import List, Optional, Any, Tuple # Добавим Any для _com_object
import time # Для возможных задержек

from .types_ import exceptions

class GuiTree:
    """
    Represents a SAP GUI Tree element (GuiShell subtype Tree).
    Provides methods to interact with tree nodes and columns.
    """
    EXPECTED_COM_TYPES = ["GuiShell", "GuiTreeControl"]

    def __init__(self, session_handle: win32com.client.CDispatch, element_id: str):
        """
        Initializes the GuiTree object.

        Args:
            session_handle: The SAP GUI session handle.
            element_id: The ID of the GuiTree element.

        Raises:
            exceptions.ElementNotFoundException: If the element is not found.
            exceptions.InvalidElementTypeException: If the element is not a GuiTree compatible type.
            exceptions.SapGuiComException: For other COM errors.
        """
        self.session_handle = session_handle # Сохраняем для возможного использования в fallback double_click
        self.element_id = element_id
        try:
            self._com_object: Any = self.session_handle.findById(self.element_id) # Указываем тип Any для COM-объекта
            self._element_type = getattr(self._com_object, "Type", "")

            if self._element_type not in self.EXPECTED_COM_TYPES:
                subtype = getattr(self._com_object, "SubType", "")
                if subtype != "Tree":
                    raise exceptions.InvalidElementTypeException(
                        f"Element '{element_id}' is Type '{self._element_type}' SubType '{subtype}', expected Tree.")
            # Дополнительная проверка на наличие ключевых методов дерева
            # (можно выбрать несколько характерных)
            if not all(hasattr(self._com_object, method) for method in ["GetAllNodeKeys", "SelectNode", "ExpandNode"]):
                 raise exceptions.InvalidElementTypeException(
                     f"Element '{element_id}' (Type: {self._element_type}) "
                     f"does not appear to be a fully functional Tree (missing one or more key methods).")

        except exceptions.InvalidElementTypeException:
            raise
        except Exception as e:
            if "findById" in str(e) or "control could not be found" in str(e).lower():
                raise exceptions.ElementNotFoundException(f"GuiTree element '{self.element_id}' not found: {e}")
            else:
                raise exceptions.SapGuiComException(f"Error initializing GuiTree for element '{self.element_id}': {e}")

    def expand_node(self, node_key: str) -> None:
        """
        Expands the specified node.

        Args:
            node_key: The key of the node to expand.
        Raises:
            exceptions.ActionException: If an error occurs during the action.
        """
        try:
            self._com_object.ExpandNode(node_key) # Используем имя метода из API
        except Exception as e:
            raise exceptions.ActionException(f"Error expanding node '{node_key}' in tree '{self.element_id}': {e}")

    def collapse_node(self, node_key: str) -> None:
        """
        Collapses the specified node.

        Args:
            node_key: The key of the node to collapse.
        Raises:
            exceptions.ActionException: If an error occurs during the action.
        """
        try:
            self._com_object.CollapseNode(node_key) # Используем имя метода из API
        except Exception as e:
            raise exceptions.ActionException(f"Error collapsing node '{node_key}' in tree '{self.element_id}': {e}")

    def select_node(self, node_key: str, ensure_visible_first: bool = False, top_node_key_if_needed: Optional[str] = None) -> None:
        """
        Selects the specified node.

        Args:
            node_key: The key of the node to select.
            ensure_visible_first (bool): If True, will attempt to set TopNode to make the node
                                         visible before selecting. Requires `top_node_key_if_needed`.
            top_node_key_if_needed (Optional[str]): The key of the node to set as TopNode if ensure_visible_first is True.
                                                    Often this can be the `node_key` itself or a parent.
        Raises:
            exceptions.ActionException: If an error occurs during the action.
        """
        try:
            if ensure_visible_first:
                if top_node_key_if_needed:
                    self.set_top_node(top_node_key_if_needed)
                    time.sleep(0.3) # Give GUI time to adjust after setting TopNode
                else:
                    # По умолчанию пытаемся сделать сам узел верхним, если он далеко
                    # Проверка, видим ли узел (это сложно без дополнительных методов API)
                    # Проще попытаться сделать его TopNode
                    self.set_top_node(node_key)
                    time.sleep(0.3)

            self._com_object.SelectNode(node_key) # Используем имя метода из API
        except Exception as e:
            raise exceptions.ActionException(f"Error selecting node '{node_key}' in tree '{self.element_id}': {e}")

    @property
    def selected_node(self) -> Optional[str]:
        """
        Gets the key of the currently selected node.
        Returns None if no node is selected or an error occurs.
        """
        try:
            return self._com_object.SelectedNode
        except Exception as e:
            # Логируем, но не прерываем, если свойство просто не установлено
            # print(f"Warning: Could not retrieve SelectedNode for tree '{self.element_id}': {e}")
            return None # Или можно перебрасывать exceptions.PropertyNotFoundException

    @property
    def top_node(self) -> Optional[str]:
        """
        Gets the key of the topmost visible node in the tree.
        Returns None if the property cannot be read.
        """
        try:
            return self._com_object.TopNode
        except Exception as e:
            # print(f"Warning: Could not retrieve TopNode for tree '{self.element_id}': {e}")
            return None # Или можно перебрасывать exceptions.PropertyNotFoundException

    def set_top_node(self, node_key: str) -> None:
        """
        Sets the topmost visible node in the tree.

        Args:
            node_key: The key of the node to set as the top node.
        Raises:
            exceptions.ActionException: If an error occurs during the action.
        """
        try:
            self._com_object.TopNode = node_key
        except Exception as e:
            raise exceptions.ActionException(f"Error setting TopNode to '{node_key}' for tree '{self.element_id}': {e}")

    def get_node_text(self, node_key: str) -> str:
        """
        Gets the display text of the specified node.

        Args:
            node_key: The key of the node.
        Returns:
            The display text of the node.
        Raises:
            exceptions.ActionException: If an error occurs.
        """
        try:
            return self._com_object.GetNodeTextByKey(node_key)
        except Exception as e:
            raise exceptions.ActionException(f"Error getting text for node '{node_key}' in tree '{self.element_id}': {e}")

    def get_all_node_keys(self) -> List[str]:
        """
        Gets a list of keys for all currently loaded (visible or expandable) nodes in the tree.
        Note: This might not return keys for nodes deep within unexpanded branches.

        Returns:
            A list of node keys.
        Raises:
            exceptions.ActionException: If an error occurs.
        """
        try:
            keys_collection = self._com_object.GetAllNodeKeys()
            return [keys_collection.Item(i) for i in range(keys_collection.Count)]
        except Exception as e:
            raise exceptions.ActionException(f"Error getting all node keys for tree '{self.element_id}': {e}")

    def get_column_names(self) -> List[str]:
        """
        Gets the technical names of the columns (for Column Trees).

        Returns:
            A list of column names.
        Raises:
            exceptions.ActionException: If an error occurs or tree is not a column tree.
        """
        try:
            # Проверка на наличие метода может быть полезна, если не все деревья его поддерживают
            if not hasattr(self._com_object, "GetColumnNames"):
                raise exceptions.ActionException(f"Tree '{self.element_id}' does not support GetColumnNames (likely not a Column Tree).")
            cols_collection = self._com_object.GetColumnNames()
            return [cols_collection.Item(i) for i in range(cols_collection.Count)]
        except Exception as e:
            raise exceptions.ActionException(f"Error getting column names for tree '{self.element_id}': {e}")

    def get_item_text(self, node_key: str, column_name: str) -> str:
        """
        Gets the text of a specific item (cell) in a node row (for List/Column Trees).

        Args:
            node_key: The key of the node.
            column_name: The technical name of the column.
        Returns:
            The text of the item.
        Raises:
            exceptions.ActionException: If an error occurs.
        """
        try:
            return self._com_object.GetItemText(node_key, column_name)
        except Exception as e:
            raise exceptions.ActionException(
                f"Error getting item text for node '{node_key}', column '{column_name}' in tree '{self.element_id}': {e}")

    def double_click_node(self, node_key: str) -> None:
        """
        Performs a double-click action on the specified node,
        preferably using the direct `DoubleClickNode` COM method.

        Args:
            node_key: The key of the node to double-click.
        Raises:
            exceptions.ActionException: If an error occurs or the action cannot be performed.
        """
        try:
            if hasattr(self._com_object, "DoubleClickNode"):
                self._com_object.DoubleClickNode(node_key)
            else:
                # Fallback: Попытка Select + Enter через session_handle (если доступен)
                # Эта часть требует, чтобы session_handle был доступен GuiTree,
                # он уже сохраняется в self.session_handle в __init__
                # print(f"Warning: Tree '{self.element_id}' does not have DoubleClickNode method. Attempting Select + Enter fallback.")
                # self.select_node(node_key) # Используем наш метод select_node
                # time.sleep(0.1) # Пауза перед отправкой VKey
                # # Предполагаем, что session_handle это сессия, а не application
                # # и что главное окно всегда wnd[0]
                # self.session_handle.findById("wnd[0]").sendVKey(0)

                # Учитывая, что прямой DoubleClickNode работал, fallback пока можно закомментировать
                # или сделать его более явным с предупреждением.
                # В наших тестах DoubleClickNode был доступен.
                raise exceptions.ActionException(
                    f"Tree '{self.element_id}' does not have a direct DoubleClickNode method. "
                    "Fallback (Select+Enter) is disabled in this version or session_handle is not correctly configured for it."
                )
        except Exception as e:
            raise exceptions.ActionException(f"Error double-clicking node '{node_key}' in tree '{self.element_id}': {e}")

    # --- Методы, которые были в вашем тестовом скрипте, но отсутствовали в GuiTree ---
    # Их можно добавить сюда, если они являются общими для GuiTree

    def get_node_children_info(self, parent_node_key: str, auto_expand: bool = True) -> List[Tuple[str, str]]:
        """
        Gets information (key, text) about the direct children of a given node.

        Args:
            parent_node_key: The key of the parent node.
            auto_expand: If True, tries to expand the parent node if it's not already expanded
                         and is expandable. Defaults to True.
        Returns:
            A list of tuples, where each tuple is (child_node_key, child_node_text).
        Raises:
            exceptions.ActionException: If an error occurs.
        """
        children_info: List[Tuple[str, str]] = []
        try:
            if auto_expand:
                if hasattr(self._com_object, "IsFolderExpandable") and hasattr(self._com_object, "IsFolderExpanded"):
                    if self._com_object.IsFolderExpandable(parent_node_key) and \
                       not self._com_object.IsFolderExpanded(parent_node_key):
                        # print(f"Auto-expanding node '{parent_node_key}' to get children.")
                        self.expand_node(parent_node_key) # Используем наш метод
                        time.sleep(0.3) # Пауза после разворачивания

            if not hasattr(self._com_object, "GetSubNodesCol"):
                raise exceptions.ActionException(f"Tree '{self.element_id}' does not support GetSubNodesCol.")

            children_keys_collection = self._com_object.GetSubNodesCol(parent_node_key)
            if children_keys_collection and hasattr(children_keys_collection, "Count"):
                for i in range(children_keys_collection.Count):
                    child_key = children_keys_collection.Item(i)
                    child_text = self.get_node_text(child_key) # Используем наш метод
                    children_info.append((child_key, child_text))
            return children_info
        except Exception as e:
            raise exceptions.ActionException(f"Error getting children for node '{parent_node_key}' in tree '{self.element_id}': {e}")

    def find_node_key_by_text(self, target_text: str, case_sensitive: bool = False, search_depth: Optional[int] = None) -> Optional[str]:
        """
        Finds the key of a node by its display text.
        Can be slow on very large trees if no search_depth is specified.

        Args:
            target_text: The text of the node to find.
            case_sensitive: Whether the search should be case-sensitive. Defaults to False.
            search_depth: Optional maximum depth to search. If None, searches all loaded nodes.
                          Not yet fully implemented for depth-limited search of *all* nodes.
                          Currently relies on GetAllNodeKeys which gets loaded nodes.
        Returns:
            The key of the first matching node, or None if not found.
        Raises:
            exceptions.ActionException: If an error occurs during the search.
        """
        # print(f"Searching for node with text: '{target_text}' (case_sensitive={case_sensitive}) in tree '{self.element_id}'")
        try:
            # TODO: Реализовать более умный поиск с учетом search_depth, если GetAllNodeKeys
            # возвращает слишком много или если нужно искать в неразвернутых ветках (что сложнее).
            # Текущая реализация ищет среди всех *загруженных* ключей.
            all_keys = self.get_all_node_keys()
            for key in all_keys:
                try:
                    node_text = self.get_node_text(key)
                    if not case_sensitive:
                        if node_text.lower() == target_text.lower():
                            return key
                    else:
                        if node_text == target_text:
                            return key
                except exceptions.ActionException:
                    # Ignore if text for a specific key cannot be retrieved during search
                    pass
            return None
        except Exception as e:
            raise exceptions.ActionException(f"Error finding node by text '{target_text}' in tree '{self.element_id}': {e}")

    # Другие полезные методы можно добавить по аналогии:
    # is_node_expanded(node_key) -> bool
    # get_parent_key(node_key) -> Optional[str]
    # ... и т.д.