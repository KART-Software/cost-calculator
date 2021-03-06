import argparse
from cost_calculator import __version__
from cost_calculator import costTableToFca, fcaToBom, supplToFca
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
    if args.fca_to_bom:
        fcaDirectoryPath = args.fca_to_bom[0]
        bomFilePath = args.fca_to_bom[1]
        fcaToBom(fcaDirectoryPath, bomFilePath)
    if args.suppl_to_fca:
        supplToFca(args.suppl_to_fca)


def _parse_args(parser: argparse.ArgumentParser) -> argparse.Namespace:
    parser.add_argument("--version",
                        action="version",
                        version="%(prog)s" + __version__)
    parser.add_argument(
        "-ctf",
        "--costtable-to-fca",
        nargs=2,
        help="5 Cost Table files' directory path and FCA files' directory path."
    )
    parser.add_argument("-ftb",
                        "--fca-to-bom",
                        nargs=2,
                        help="FCA files' directory path and BOM file path.")
    parser.add_argument(
        "-stf",
        "--suppl-to-fca",
        help="FCA files' directory path to write on the link to Supplement PDF."
    )

    return parser.parse_args()


if __name__ == "__main__":
    main()
