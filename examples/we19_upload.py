"""Upload and execute an IDoc in WE19."""
from pathlib import Path

from sapscriptwizard import SapSession, WE19


def main() -> None:
    session = SapSession.from_gui()
    we19 = WE19(session)
    we19.upload_xml(Path("idoc.xml")).execute()


if __name__ == "__main__":  # pragma: no cover
    main()
