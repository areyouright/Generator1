#!/usr/bin/env python3
from __future__ import annotations

import argparse

from st_codegen import (
    read_ai_points,
    read_di_points,
    read_do_points,
    read_plc_config,
    read_protocol_params,
    read_rtd_points,
    render_protocol_excel,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Read PLC/AI/RTD/DI/DO sheets from Excel and generate protocol.xlsx"
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
        default="protocol.xlsx",
        help="Output path for generated protocol Excel",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    config = read_plc_config(args.xlsx_path)
    protocol = read_protocol_params(args.xlsx_path)
    ai_points = read_ai_points(args.xlsx_path)
    di_points = read_di_points(args.xlsx_path)
    do_points = read_do_points(args.xlsx_path)
    rtd_points = read_rtd_points(args.xlsx_path)
    render_protocol_excel(config, protocol, ai_points, di_points, do_points, rtd_points, args.output_path)
    print(f"Generated {args.output_path}")


if __name__ == "__main__":
    main()
