"""Minimal example for connecting to SAP and reading window title."""
from sapscriptwizard import SapSession

session = SapSession()
session.connect()
window = session.window()
print(window.dump_gui_structure(max_depth=1))
