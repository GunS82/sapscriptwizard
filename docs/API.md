# API Overview

This document summarizes the main classes and helpers provided by **sapscriptwizard**.
It covers public methods across modules and demonstrates common workflows such as
session management, element search, table export and parallel execution.

## Creating a `Sapscript` instance

```python
from sapscriptwizard import Sapscript
sap = Sapscript()
```

A `Sapscript` object is the entry point for launching or attaching to SAP GUI sessions.

### `launch_sap()`
Launches SAP with credentials using `sapshcut.exe` and waits for the main window. Parameters include `sid`, `client`, `user`, `password` and optional flags such as `maximise` and `language`. Raises `WindowDidNotAppearException` if the window is not detected.

### `quit()`
Attempts to log off the first session and then kills `saplogon.exe`.

### `attach_window(connection_index, session_index)`
Returns a :class:`Window` bound to the given connection and session indices.

### `open_new_window(window_to_handle_opening)`
Creates a new session using an existing :class:`Window` object.

### `start_saplogon()`
Starts `saplogon.exe` if the SAP GUI scripting object is not found. The method returns `True` when the process is running or was started successfully, `False` otherwise. It checks for the executable and waits until the scripting object becomes available【F:sapscriptwizard.py†L343-L391】.

### Connection helpers
* `get_connection_count()` – number of SAP connections.
* `get_connection_info(index)` – description of a connection.
* `get_active_session_indices(index)` – list of active session numbers.
* `get_session_info(conn_idx, sess_idx)` – information about a particular session.
* `get_all_connections_info()` – combined information for all connections.
* `find_session_by_sid_user(sid, user)` – search for a session with matching SID and user.

### Screenshots and history
* `enable_screenshots_on_error()` / `disable_screenshots_on_error()` – toggle automatic capture when `handle_exception_with_screenshot()` is used.
* `set_screenshot_directory(path)` – directory for saving screenshots.
* `handle_exception_with_screenshot(exc)` – logs the exception and optionally saves a screenshot.
* `disable_history()` / `enable_history()` – toggle the SAP GUI input history.

## `Window`
Returned by `Sapscript.attach_window`, this class wraps a SAP session handle.
Key methods include:

* `maximize()`, `restore()`, `close_window()` – basic window management.
* `navigate(action)` – press Enter/back/end/cancel/save using :class:`NavigateAction`.
* `start_transaction(tcode)` – send a transaction code.
* `press(element)`, `select(element)` – interact with controls by ID.
* `write(element, text)` and `read(element)` – modify or read `text` property.
* `set_checkbox(element, bool)` / `is_selected(element)` – interact with check boxes.
* `visualize(element, seconds=1)` – draw a red frame around an element for debugging.
* `exists(element)` – check if an element is present.
* `send_v_key(value)` – send a virtual key to the window or element.
* `read_html_viewer(element)` – return HTML content from a `GuiHTMLViewer`.
* `read_shell_table(element)` – returns a :class:`ShellTable` instance for ALV grids.
* `get_status_message()` and `assert_status_bar()` – read and verify the status bar.
* `select_menu_item_by_name(menu_path)` – choose items by text from a menu hierarchy.
* `start_transaction_robust(tcode)` – start a transaction with status bar checks.
* Utility methods for scanning GUI structure (`dump_element_state`, `save_gui_snapshot`, etc.).
* Locator-based helpers (`press_by_locator`, `write_by_locator`, etc.) resolve semantic descriptions using :class:`sapscriptwizard_semantic.element_finder.SapElementFinder`.

## `ShellTable`
Represents a GuiShell/GuiGridView table and exposes data via `polars` or `pandas`.

* Constructed with a session handle and element ID. If `load_table=True`, the table is scrolled to load all rows.
* Supports index access and iteration (`table[0]`, `for row in table`).
* `to_polars_dataframe()` / `to_pandas_dataframe()` – return the underlying data frame.
* `to_dict()` / `to_dicts()` – convert table data to dictionaries.
* `cell(row, column)` – access a single cell by index or column name.
* `load()` – scrolls the table to ensure all rows are loaded.
* `press_button(button)` – press a toolbar button inside the shell table.
* `select_rows([0,1])` – select rows by 0-based index.
* `change_checkbox(row, column_id, flag)` – toggle a checkbox in a cell.
* `to_csv(path)` – save the table to CSV.

## `GuiTree`
Wrapper for tree controls. Instantiate via `window.get_tree(element_id)`.

* `expand_node(key)` / `collapse_node(key)` – show or hide a node.
* `select_node(key, ensure_visible_first=False)` – select a node, optionally scrolling it into view.
* `selected_node` and `top_node` properties – read current selection and top visible node.
* `set_top_node(key)` – scroll the tree so the node is at the top.
* `get_node_text(key)` – display text of the node.
* `get_all_node_keys()` – list of loaded node keys.
* `get_column_names()` – names of columns for column trees.
* `get_item_text(node, column)` – read a value from a node row.
* `double_click_node(key)` – perform a double click.
* Additional helpers include `get_node_children_info()` and `find_node_key_by_text()`.

## Parallel execution (`parallel.api`)
Use `run_parallel` to perform a worker function across multiple sessions. It can open new windows or reuse existing ones.

```python
from sapscriptwizard.parallel import run_parallel
from sapscriptwizard import Window

def worker(window: Window, data: list):
    # interact with SAP using the provided Window
    pass

run_parallel(
    enabled=True,
    num_processes=2,
    worker_function=worker,
    input_data_list=["item1", "item2"],
    interactive=True
)
```

### `run_parallel()` parameters
* `enabled` – run sequentially (`False`) or in parallel (`True`).
* `num_processes` – number of worker processes when opening new sessions.
* `worker_function` – callback receiving a `Window` and a list of data.
* `input_data_list` / `input_data_file` – data to process.
* `interactive` – prompt for connection and sessions when `True`.
* Additional `runner_kwargs` are passed to :class:`SapParallelRunner`.

## Example workflow

```python
from sapscriptwizard import Sapscript

# Start SAP and attach to the first session
sap = Sapscript()
if not Sapscript.start_saplogon():
    raise RuntimeError('SAP Logon not found')
window = sap.attach_window(0, 0)

# Open a new session
sap.open_new_window(window)
new_window = sap.attach_window(0, 1)

# Locate elements by semantics
new_window.write_by_locator('@ Пользователь', 'DEMO')
new_window.write_by_locator('@ Пароль', 'secret')
new_window.press_by_locator('=Вход')

# Work with a tree
tree = new_window.get_tree('wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell')
keys = tree.get_all_node_keys()
if keys:
    tree.double_click_node(keys[0])

# Export an ALV grid
shell = new_window.read_shell_table('wnd[0]/usr/cntlGRID_CONTAINER/shellcont/shell')
shell.to_csv('output.csv')
```

The basic usage snippet above mirrors the one in the project `README.md` and can be extended with the code shown here for more advanced automation.
