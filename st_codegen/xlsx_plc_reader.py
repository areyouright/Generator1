from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re
import xml.etree.ElementTree as ET
import zipfile

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"m": MAIN_NS, "r": REL_NS, "pr": PKG_REL_NS}


@dataclass(frozen=True)
class PlcConfig:
    count_ai: int
    count_rtd: int
    count_di: int
    count_do: int
    count_lines: int
    count_dt: int
    first_group: int

    @property
    def count_modules(self) -> int:
        return self.count_ai + self.count_rtd + self.count_di + self.count_do


@dataclass(frozen=True)
class RtdPoint:
    tag: str
    module: int
    channel: int


@dataclass(frozen=True)
class PhasePoint:
    line_no: int
    phase_code: int


@dataclass(frozen=True)
class AiPoint:
    tag: str
    module: int
    channel: int


@dataclass(frozen=True)
class DoPoint:
    tag: str
    module: int
    channel: int


@dataclass(frozen=True)
class DiPoint:
    tag: str
    module: int
    channel: int


def _col_to_index(col: str) -> int:
    result = 0
    for ch in col:
        result = result * 26 + (ord(ch) - 64)
    return result


def _read_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    strings: list[str] = []
    for si in root.findall("m:si", NS):
        text = "".join(t.text or "" for t in si.findall(".//m:t", NS))
        strings.append(text)
    return strings


def _resolve_workbook_targets(zf: zipfile.ZipFile) -> dict[str, str]:
    rels_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    targets: dict[str, str] = {}
    for rel in rels_root.findall("pr:Relationship", NS):
        rid = rel.attrib["Id"]
        target = rel.attrib["Target"]
        targets[rid] = target
    return targets


def _sheet_path_by_name(zf: zipfile.ZipFile, sheet_name: str) -> str:
    wb_root = ET.fromstring(zf.read("xl/workbook.xml"))
    rid_to_target = _resolve_workbook_targets(zf)

    for sheet in wb_root.findall("m:sheets/m:sheet", NS):
        if sheet.attrib.get("name") != sheet_name:
            continue
        rid = sheet.attrib[f"{{{REL_NS}}}id"]
        target = rid_to_target[rid]
        if target.startswith("/"):
            return target.lstrip("/")
        if target.startswith("xl/"):
            return target
        return f"xl/{target}"

    raise ValueError(f"Sheet '{sheet_name}' not found in workbook")


def _read_sheet_rows(zf: zipfile.ZipFile, sheet_path: str, shared_strings: list[str]) -> list[dict[int, str]]:
    root = ET.fromstring(zf.read(sheet_path))
    rows: list[dict[int, str]] = []

    for row in root.findall(".//m:sheetData/m:row", NS):
        row_data: dict[int, str] = {}
        for cell in row.findall("m:c", NS):
            ref = cell.attrib.get("r", "")
            match = re.match(r"([A-Z]+)", ref)
            if not match:
                continue
            col = _col_to_index(match.group(1))
            cell_type = cell.attrib.get("t")
            value_node = cell.find("m:v", NS)

            if value_node is None:
                inline = cell.find("m:is", NS)
                if inline is None:
                    continue
                value = "".join(t.text or "" for t in inline.findall(".//m:t", NS))
            else:
                raw = value_node.text or ""
                value = shared_strings[int(raw)] if cell_type == "s" else raw

            row_data[col] = value.strip()

        if row_data:
            rows.append(row_data)

    return rows


def _load_sheet_rows(xlsx_path: str | Path, sheet_name: str) -> list[dict[int, str]]:
    xlsx = Path(xlsx_path)
    with zipfile.ZipFile(xlsx) as zf:
        shared_strings = _read_shared_strings(zf)
        sheet_path = _sheet_path_by_name(zf, sheet_name)
        return _read_sheet_rows(zf, sheet_path, shared_strings)


def read_plc_config(xlsx_path: str | Path) -> PlcConfig:
    rows = _load_sheet_rows(xlsx_path, "PLC")

    params: dict[str, int] = {}
    for row in rows:
        key = row.get(2, "")
        value = row.get(1, "")
        if not key.startswith("="):
            continue
        key = key[1:].strip().lower()
        if value == "":
            continue
        try:
            params[key] = int(float(value))
        except ValueError:
            continue

    required = {
        "count ai",
        "count rtd",
        "count di",
        "count do",
        "count lines",
        "count dt",
        "first_group",
    }

    missing = [excel_key for excel_key in sorted(required) if excel_key not in params]
    if missing:
        raise ValueError(f"Missing PLC parameters in Excel: {', '.join(missing)}")

    return PlcConfig(
        count_ai=params["count ai"],
        count_rtd=params["count rtd"],
        count_di=params["count di"],
        count_do=params["count do"],
        count_lines=params["count lines"],
        count_dt=params["count dt"],
        first_group=params["first_group"],
    )


def read_rtd_points(xlsx_path: str | Path) -> list[RtdPoint]:
    rows = _load_sheet_rows(xlsx_path, "RTD")
    points: list[RtdPoint] = []

    for row in rows[1:]:
        tag = row.get(1, "")
        module = row.get(2, "")
        channel = row.get(3, "")

        if not tag or not module or not channel:
            continue

        try:
            points.append(
                RtdPoint(
                    tag=tag,
                    module=int(float(module)),
                    channel=int(float(channel)),
                )
            )
        except ValueError:
            continue

    return points


def _phase_code_from_text(phase_text: str) -> int:
    normalized = phase_text.upper().replace(" ", "")
    normalized = normalized.replace("A", "А").replace("B", "В").replace("C", "С")
    parts = [p for p in re.split(r"[^АВС]+", normalized) if p]
    phase_set = set(parts)

    mapping = {
        frozenset({"А"}): 1,
        frozenset({"В"}): 2,
        frozenset({"А", "В"}): 3,
        frozenset({"С"}): 4,
        frozenset({"А", "С"}): 5,
        frozenset({"В", "С"}): 6,
        frozenset({"А", "В", "С"}): 7,
    }
    key = frozenset(phase_set)
    if key not in mapping:
        raise ValueError(f"Unsupported phase value: '{phase_text}'")
    return mapping[key]


def read_phase_points(xlsx_path: str | Path) -> list[PhasePoint]:
    rows = _load_sheet_rows(xlsx_path, "Phase")
    points: list[PhasePoint] = []

    for row in rows[1:]:
        phase_text = row.get(1, "")
        line_text = row.get(2, "")

        if not phase_text or not line_text:
            continue

        match = re.search(r"(\d+)", line_text)
        if not match:
            continue

        points.append(
            PhasePoint(
                line_no=int(match.group(1)),
                phase_code=_phase_code_from_text(phase_text),
            )
        )

    return sorted(points, key=lambda p: p.line_no)


def read_ai_points(xlsx_path: str | Path) -> list[AiPoint]:
    rows = _load_sheet_rows(xlsx_path, "AI")
    points: list[AiPoint] = []

    for row in rows[1:]:
        tag = row.get(1, "")
        module = row.get(2, "")
        channel = row.get(3, "")

        if not tag or not module or not channel:
            continue

        try:
            points.append(
                AiPoint(
                    tag=tag,
                    module=int(float(module)),
                    channel=int(float(channel)),
                )
            )
        except ValueError:
            continue

    return points


def read_do_points(xlsx_path: str | Path) -> list[DoPoint]:
    rows = _load_sheet_rows(xlsx_path, "DO")
    points: list[DoPoint] = []

    for row in rows[1:]:
        tag = row.get(1, "")
        module = row.get(2, "")
        channel = row.get(3, "")

        if not tag or not module or not channel:
            continue

        try:
            points.append(
                DoPoint(
                    tag=tag,
                    module=int(float(module)),
                    channel=int(float(channel)),
                )
            )
        except ValueError:
            continue

    return points


def read_di_points(xlsx_path: str | Path) -> list[DiPoint]:
    rows = _load_sheet_rows(xlsx_path, "DI")
    points: list[DiPoint] = []

    for row in rows[1:]:
        tag = row.get(1, "")
        module = row.get(2, "")
        channel = row.get(3, "")

        if not tag or not module or not channel:
            continue

        try:
            points.append(
                DiPoint(
                    tag=tag,
                    module=int(float(module)),
                    channel=int(float(channel)),
                )
            )
        except ValueError:
            continue

    return points
