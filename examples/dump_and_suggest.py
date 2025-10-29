"""Dump GUI structure and suggest locators."""
from sapscriptwizard import SapSession, explain_id, suggest_locator

session = SapSession()
session.connect()
window = session.window()
structure = window.dump_gui_structure(max_depth=2)
# Example element id placeholder
info = explain_id(structure, structure["window_id"])
print([hint.describe() for hint in suggest_locator(info)])
