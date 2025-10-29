"""Command line interface for sapscriptwizard."""
from __future__ import annotations

import argparse
import json
import logging
from typing import Any

from ..core.session import SapSession

log = logging.getLogger(__name__)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="sapwiz", description="Inspect SAP GUI windows")
    parser.add_argument("--dump", action="store_true", help="Dump GUI structure as JSON")
    parser.add_argument("--window-index", type=int, default=0, help="Window index to inspect")
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    logging.basicConfig(level=logging.INFO)

    session = SapSession()
    session.connect()
    window = session.window(args.window_index)

    if args.dump:
        payload: dict[str, Any] = window.dump_gui_structure()
        print(json.dumps(payload, indent=2, ensure_ascii=False))
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
