# VBS to Python reference

| VBS statement | Python equivalent |
|---------------|------------------|
| `session.findById("wnd[0]/usr/txtSOME").text = "42"` | `session.window().write("usr/txtSOME", "42")` |
| `session.findById("wnd[0]/usr/btnEXECUTE").press()` | `session.window().press("usr/btnEXECUTE")` |
| `session.findById("wnd[0]/usr/cntlGRID").GetCellValue(0, "STATUS")` | `ShellTable(session.window(), "usr/cntlGRID").to_dicts()[0]["STATUS"]` |
| `session.findById("wnd[0]").sendVKey(0)` | `session.window().send_vkey(0)` |
