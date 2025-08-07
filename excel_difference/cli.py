#!/usr/bin/env python3
"""
Command-line interface for Excel difference generator.
"""

import argparse
import sys
from pathlib import Path

from excel_difference.excel_diff import excel_diff


def main():
    """Main CLI function."""
    parser = argparse.ArgumentParser(
        description="Generate a difference Excel file between two input Excel files."
    )
    parser.add_argument("file1", help="Path to the first Excel file")
    parser.add_argument("file2", help="Path to the second Excel file")
    parser.add_argument("output", help="Path for the output Excel file")
    parser.add_argument(
        "--key-column",
        type=int,
        default=1,
        help="Column to use for matching rows (1-based, default=1)",
    )

    args = parser.parse_args()

    # Check if input files exist
    if not Path(args.file1).exists():
        print(f"Error: File '{args.file1}' does not exist.")
        sys.exit(1)

    if not Path(args.file2).exists():
        print(f"Error: File '{args.file2}' does not exist.")
        sys.exit(1)

    try:
        excel_diff(args.file1, args.file2, args.output, args.key_column)
        print(f"Successfully generated difference file: {args.output}")
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
