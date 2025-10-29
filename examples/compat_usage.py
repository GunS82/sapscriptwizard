"""Illustrate the legacy compatibility layer."""
import warnings

from sapscriptwizard import SapSession
from sapscriptwizard import shell_table

warnings.simplefilter("default", DeprecationWarning)


def main() -> None:
    session = SapSession.from_gui()
    table = shell_table.ShellTable(session.raw, "usr/cntlGRID")
    print(table.to_dicts())


if __name__ == "__main__":  # pragma: no cover
    main()
