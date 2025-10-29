# Migration guide from legacy API

| Legacy API | New API |
|------------|---------|
| `sapscriptwizard.Sapscript.launch_sap()` | `SapSession.connect()` |
| `window.Window.write(id, text)` | `Window.write(id, text)` |
| `window.Window.press(id)` | `Window.press(id)` |
| `shell_table.ShellTable.to_dicts()` | `gui.controls.table.ShellTable.to_dicts()` |
| `gui_tree.GuiTree` | `gui.controls.tree.GuiTree` |

Additional changes:

* Legacy modules now emit `DeprecationWarning` and will be removed in the next major release.
* Use `SapSession.window()` to retrieve `Window` objects instead of instantiating `Window` directly.
* Configure logging through `sapscriptwizard.configure_logging()`.
