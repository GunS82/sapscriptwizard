"""Wait helper usage."""
from sapscriptwizard import SapSession, waits


def main() -> None:
    session = SapSession.from_gui()
    window = session.window()
    waits.wait_ready(window)
    waits.wait_visible(window, "usr/btnConfirm")
    waits.wait_status(window, text="Ready")


if __name__ == "__main__":  # pragma: no cover
    main()
