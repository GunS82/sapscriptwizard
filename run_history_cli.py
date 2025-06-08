from argparse import ArgumentParser
from pathlib import Path
from typing import Any
from run_history import RunHistory


def main() -> None:
    parser = ArgumentParser(description="Show run history records")
    parser.add_argument("--db", default="run_history.db", help="Path to history database")
    parser.add_argument("--limit", type=int, default=10, help="Number of records to display")
    args = parser.parse_args()

    history = RunHistory(Path(args.db))
    runs = history.fetch_runs(args.limit)
    for run in runs:
        print(run)


if __name__ == "__main__":
    main()
