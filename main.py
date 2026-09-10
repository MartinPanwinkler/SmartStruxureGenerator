from __future__ import annotations

from config import setup_logging
from gui import run_gui


def main() -> None:
    setup_logging()
    run_gui()


if __name__ == "__main__":
    main()
