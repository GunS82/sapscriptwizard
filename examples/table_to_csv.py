"""Example converting an ALV grid into a CSV file."""
import csv
from pathlib import Path

from sapscriptwizard import SapSession, ShellTable


def main() -> None:
    session = SapSession.from_gui()
    window = session.window()
    table = ShellTable(window, "usr/cntlGRID")
    data = table.to_dicts()
    path = Path("table_export.csv")
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=data[0].keys() if data else [])
        writer.writeheader()
        writer.writerows(data)
    print(f"Exported {len(data)} rows to {path}")


if __name__ == "__main__":  # pragma: no cover
    main()
