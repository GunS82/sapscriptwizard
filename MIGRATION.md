# Migration guide from 0.x to 1.0

| Legacy call | New API |
|-------------|---------|
| `window.Window(...).press(element)` | `session.window().press(element)` |
| `session.findById(...).text = "X"` | `session.window().write(element, "X")` |
| `shell_table.ShellTable(session, id)` | `ShellTable(session.window(), id)` |
| `SapElementFinder(...).find(query)` | `window.finder.find(query)` |
| `GuiTree(session, id).select(...)` | `GuiTree(session.window(), id).select_path(...)` |
