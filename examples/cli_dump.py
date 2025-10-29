"""Invoke the CLI from Python."""
from sapscriptwizard.cli.sapwiz import main


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main(["dump"]))
