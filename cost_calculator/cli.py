import argparse
from typing import List, Optional


def main():
    parser = argparse.ArgumentParser(description=main.__doc__)
    args = _parse_args(parser)


def _parse_args(
    parser: argparse.ArgumentParser, args: Optional[List] = None
) -> argparse.Namespace:
    parser.add_argument(
        "FCA File",
        "Table Materials",
        "Table Processes",
        "Table Process Multipliers",
        "Table Fasteners",
        "Table Tooling",
        help="The FCA files' path, Cost Table file's path",
        nargs="?",
    )
