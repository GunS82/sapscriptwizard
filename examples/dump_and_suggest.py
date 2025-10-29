"""Dump GUI structure and print locator suggestions."""
from sapscriptwizard import SapSession, explain_id, suggest_locator


def main() -> None:
    session = SapSession.from_gui()
    window = session.window()
    structure = window.dump_gui_structure()
    print(structure)
    element_id = "usr/txtExample"
    print(explain_id(window, element_id).to_dict())
    print([hint.to_dict() for hint in suggest_locator(window, element_id)])


if __name__ == "__main__":  # pragma: no cover
    main()
