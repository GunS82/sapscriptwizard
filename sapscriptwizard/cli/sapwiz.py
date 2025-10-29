"""Command line interface for the SAP Script Wizard."""
from __future__ import annotations

import argparse
import json
import logging
import sys

from ..core.session import SapSession
from ..helpers.explain import explain_id, suggest_locator


def _configure_logging(verbose: bool) -> None:
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format='{"time":"%(asctime)s","level":"%(levelname)s","message":"%(message)s"}',
    )


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(prog="sapwiz", description="SAP Script Wizard CLI")
    parser.add_argument("command", choices=["dump", "explain", "suggest"], help="Action to execute")
    parser.add_argument("argument", nargs="?", help="Element ID for explain/suggest")
    parser.add_argument("--verbose", action="store_true", help="Enable debug logging")
    parser.add_argument("--max-depth", type=int, default=2)
    parser.add_argument("--max-children", type=int, default=100)
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    _configure_logging(args.verbose)
    session = SapSession.from_gui()
    window = session.window()
    if args.command == "dump":
        dump = window.dump_gui_structure(max_depth=args.max_depth, max_children=args.max_children)
        print(json.dumps(dump, indent=2, ensure_ascii=False))
        return 0
    if args.argument is None:
        raise SystemExit("argument is required for this command")
    if args.command == "explain":
        info = explain_id(window, args.argument)
        print(json.dumps(info.to_dict(), indent=2, ensure_ascii=False))
        return 0
    if args.command == "suggest":
        hints = suggest_locator(window, args.argument)
        print(json.dumps([hint.to_dict() for hint in hints], indent=2, ensure_ascii=False))
        return 0
    raise SystemExit(f"Unsupported command {args.command}")


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
