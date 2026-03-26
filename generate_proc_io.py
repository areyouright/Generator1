#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path

from st_codegen import read_phase_points, read_plc_config, read_rtd_points, render_proc_io


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Read PLC+RTD+Phase sheets from Excel and generate Proc_IO.txt"
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
        default="Proc_IO.txt",
        help="Output path for generated Proc_IO text",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    config = read_plc_config(args.xlsx_path)
    rtd_points = read_rtd_points(args.xlsx_path)
    phase_points = read_phase_points(args.xlsx_path)
    proc_io_text = render_proc_io(config, rtd_points, phase_points)
    Path(args.output_path).write_text(proc_io_text, encoding="utf-8")
    print(f"Generated {args.output_path}")


if __name__ == "__main__":
    main()
