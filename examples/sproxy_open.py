"""Open a service proxy."""
from sapscriptwizard import SapSession, SProxy


def main() -> None:
    session = SapSession.from_gui()
    sproxy = SProxy(session)
    sproxy.open_service("ZCO_SERVICE")


if __name__ == "__main__":  # pragma: no cover
    main()
