"""Read program source via SE80 feature wrapper."""
from sapscriptwizard import SapSession, SE80


def main() -> None:
    session = SapSession.from_gui()
    se80 = SE80(session)
    print(se80.read_source("ZREPORT"))


if __name__ == "__main__":  # pragma: no cover
    main()
