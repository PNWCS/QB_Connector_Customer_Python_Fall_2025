"""Command-line interface for the customers synchroniser."""

from __future__ import annotations

import argparse
import sys

from .runner import run_customer_sync


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Synchronise customers between Excel and QuickBooks"
    )
    parser.add_argument(
        "--workbook",
        required=True,
        help="Excel workbook containing the customers worksheet",
    )
    parser.add_argument("--output", help="Optional JSON output path")

    args = parser.parse_args(argv)

    path = run_customer_sync("", args.workbook, output_path=args.output)
    print(f"Report written to {path}")
    return 0


if __name__ == "__main__":  # pragma: no cover - manual invocation
    sys.exit(main())