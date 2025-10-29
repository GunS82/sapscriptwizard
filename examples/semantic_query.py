"""Resolve semantic locators."""
from sapscriptwizard import SapSession


def main() -> None:
    session = SapSession.from_gui()
    window = session.window()
    print(window.finder.find('@"User"'))
    print(window.finder.find('="Execute"'))


if __name__ == "__main__":  # pragma: no cover
    main()
