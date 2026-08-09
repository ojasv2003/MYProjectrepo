#!/usr/bin/env python3
"""Quick terminal runner for the Logistics Data Agent."""

from __future__ import annotations

import argparse
import json
import logging
import os
import sys

from dotenv import load_dotenv

from agents.master_agent import MasterAgent


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Collect Indian MMLH / MMLP logistics hub data into the master CSV. "
            "Pass one or more URLs, or use --search to discover URLs via DuckDuckGo."
        )
    )
    parser.add_argument(
        "urls",
        nargs="*",
        help="One or more public page URLs to fetch and extract",
    )
    parser.add_argument(
        "--search",
        "-s",
        metavar="QUERY",
        help='Search DuckDuckGo then process results, e.g. "MMLP Nagpur logistics park capacity"',
    )
    parser.add_argument(
        "--max-urls",
        type=int,
        default=8,
        help="Maximum URLs to process from search (default: 8)",
    )
    parser.add_argument(
        "--pretty",
        action="store_true",
        help="Pretty-print JSON report",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    load_dotenv()
    logging.basicConfig(
        level=os.getenv("LOG_LEVEL", "INFO").upper(),
        format="%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
    )

    parser = build_parser()
    args = parser.parse_args(argv)

    if not args.search and not args.urls:
        parser.print_help()
        print(
            "\nExamples:\n"
            '  python run_cli.py --search "MMLP Nagpur logistics park capacity"\n'
            "  python run_cli.py https://nhlml.org/multi-modal-logistics-park\n",
            file=sys.stderr,
        )
        return 2

    master = MasterAgent()

    if args.search:
        report = master.run_search(args.search, max_urls=args.max_urls)
    else:
        report = master.run_urls(args.urls)

    indent = 2 if args.pretty else None
    print(json.dumps(report, indent=indent, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
