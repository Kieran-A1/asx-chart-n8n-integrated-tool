from __future__ import annotations

import argparse
import json
from pathlib import Path

from .pipeline import run_asx_report


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Generate a Yahoo Finance chart report (image + docx + pdf) and optionally email "
            "the PDF with Apple Mail."
        )
    )
    parser.add_argument("--asx-code", default="", help="Ticker code (example: BHP).")
    parser.add_argument(
        "--recipient",
        default="test@gmail.com",
        help="Recipient email address.",
    )
    parser.add_argument(
        "--output-dir",
        default="output",
        help="Output folder for generated files.",
    )
    parser.add_argument(
        "--email-subject",
        default="",
        help="Optional custom email subject.",
    )
    parser.add_argument(
        "--email-body",
        default="",
        help="Optional custom email body.",
    )
    parser.add_argument(
        "--no-email",
        action="store_true",
        help="Generate files but skip the Mail app send step.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    result = run_asx_report(
        asx_code=args.asx_code or None,
        recipient=args.recipient,
        output_dir=Path(args.output_dir),
        email_subject=args.email_subject or None,
        email_body=args.email_body or None,
        send_email=not args.no_email,
    )
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
