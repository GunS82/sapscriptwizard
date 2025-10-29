"""Minimal example that reads a field and presses a button."""
from sapscriptwizard import SapSession


def main() -> None:
    session = SapSession.from_gui()
    window = session.window()
    print(window.read_text("usr/txtUser"))
    window.press("usr/btnConfirm")


if __name__ == "__main__":  # pragma: no cover
    main()
