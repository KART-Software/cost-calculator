import argparse
from cost_calculator import costTableToFca
from typing import List, Optional


def main():
    """Command line application of cost calculator"""
    parser = argparse.ArgumentParser(description=main.__doc__)
    args = _parse_args(parser)
    _perform_args(args)


def _perform_args(args: argparse.Namespace) -> None:
    if args.costtable_to_fca:
        costTableDirectoryPath = args.costtable_to_fca[0]
        fcaDirectoryPath = args.costtable_to_fca[1]
        costTableToFca(costTableDirectoryPath, fcaDirectoryPath)


def _parse_args(parser: argparse.ArgumentParser) -> argparse.Namespace:
    # parser.add_argument(
    #     "Path of directory containing 5 Cost Table files",
    #     "Path of directory containing FCA files",
    #     help="The FCA files' path, Cost Table file's path",
    #     nargs="?",
    # )
    parser.add_argument(
        "-ctf",
        "--costtable-to-fca",
        nargs=2,
        help="5 Cost Table files' directory path and FCA files' directory path"
    )
    parser.add_argument("-ftb",
                        "--fca-to-bom",
                        nargs=2,
                        help="FCA files' directory path and BOM file path")

    return parser.parse_args()


if __name__ == "__main__":
    main()
