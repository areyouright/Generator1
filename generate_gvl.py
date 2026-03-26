#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path

from st_codegen import read_plc_config, render_gvl


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Read PLC sheet from Excel and generate GVL.txt"
    )
    parser.add_argument(
        "xlsx_path",
        nargs="?",
        default="input_softitek.xlsx",
        help="Path to source Excel file (.xlsx)",
    )
    parser.add_argument(
        "output_path",
        nargs="?",
        default="GVL.txt",
        help="Output path for generated GVL text",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    config = read_plc_config(args.xlsx_path)
    gvl_text = render_gvl(config)
    Path(args.output_path).write_text(gvl_text, encoding="utf-8")
    print(f"Generated {args.output_path}")


if __name__ == "__main__":
    main()
