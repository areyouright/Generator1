from __future__ import annotations

import html
import re
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable

from .xlsx_plc_reader import AiPoint, DiPoint, DoPoint, PlcConfig, RtdPoint


AI_HEADERS = [
    "№",
    "Параметр",
    "Модуль",
    "Канал",
    "Клеммник",
    "Контакт",
    "Цепь",
    "Соответствие на мнемосхеме",
    "Результат проверки канала",
]

DI_HEADERS = [
    "№",
    "Параметр",
    "Модуль",
    "Канал",
    "Реле",
    "Контакт",
    "Цепь",
    "Соответствие на мнемосхеме",
    "Результат проверки канала",
]

DO_HEADERS = [
    "№",
    "Параметр",
    "Модуль",
    "Канал",
    "Цепь",
    "Соответствие на мнемосхеме",
    "Результат проверки канала",
]


def _phase_letter(current_no: int) -> str:
    return {1: "A", 2: "B", 3: "C"}.get(current_no, "")


def _ai_parameter_name(tag: str) -> str:
    normalized = tag.strip().lower()
    ta_match = re.fullmatch(r"ta(\d+)\.(\d+)", normalized)
    if ta_match:
        line_no = int(ta_match.group(1))
        current_no = int(ta_match.group(2))
        phase = _phase_letter(current_no)
        return f"Ток линии QFD{line_no} фаза {phase}" if phase else f"Ток линии QFD{line_no}"

    tan_match = re.fullmatch(r"tan(\d+)", normalized)
    if tan_match:
        return f"Ток утечки линии QFD{int(tan_match.group(1))}"

    return tag.upper()


def _di_parameter_name(tag: str) -> str:
    normalized = tag.strip().lower()
    if re.fullmatch(r"qfds\d+", normalized):
        return f"Состояние {tag.upper()}"
    if re.fullmatch(r"qfd\d+", normalized):
        return f"Авария {tag.upper()}"
    if re.fullmatch(r"km\d+(?:\.\d+)?", normalized):
        return f"Вкл/откл линии №{normalized[2:]}"
    if normalized in {"manu", "manuon"}:
        return "Обогрев в ручном режиме"
    if normalized == "auto":
        return "Обогрев в автоматическом режиме"
    return f"Сигнал {tag.upper()}"


def _do_parameter_name(tag: str) -> str:
    normalized = tag.strip().lower()
    if re.fullmatch(r"km\d+(?:\.\d+)?", normalized):
        return f"Вкл/откл линии №{normalized[2:]}"
    if normalized == "lineson":
        return "Обогрев включен"
    if normalized == "alall":
        return "Общая авария"
    if normalized in {"aldt", "alsensor", "failsensor"}:
        return "Авария датчика температуры"
    return f"Команда {tag.upper()}"


def _collect_rows(
    config: PlcConfig,
    ai_points: list[AiPoint],
    di_points: list[DiPoint],
    do_points: list[DoPoint],
    rtd_points: list[RtdPoint],
) -> tuple[list[list[object]], list[list[object]], list[list[object]]]:
    ai_rows: list[list[object]] = [AI_HEADERS]
    index = 1
    for point in ai_points:
        ai_rows.append(
            [
                index,
                _ai_parameter_name(point.tag),
                f"A3.{point.module}",
                point.channel,
                point.tag.upper(),
                1,
                "24В,4..20мА",
                "соотв.",
                "исправен",
            ]
        )
        index += 1

    for dt_index, point in enumerate(rtd_points[: config.count_dt + 1]):
        ai_rows.append(
            [
                index,
                f"Датчик температуры {dt_index}",
                f"A4.{point.module}",
                point.channel,
                point.tag.upper(),
                "",
                "термосопротивление",
                "соотв.",
                "исправен",
            ]
        )
        index += 1

    di_rows: list[list[object]] = [DI_HEADERS]
    for i, point in enumerate(di_points, start=1):
        di_rows.append(
            [
                i,
                _di_parameter_name(point.tag),
                f"A1.{point.module}",
                point.channel,
                point.tag.upper(),
                "",
                "24В, вх.",
                "соответств.",
                "исправен",
            ]
        )

    do_rows: list[list[object]] = [DO_HEADERS]
    for i, point in enumerate(do_points, start=1):
        do_rows.append(
            [
                i,
                _do_parameter_name(point.tag),
                f"A2.{point.module}",
                point.channel,
                "24В, вых.",
                "соответств.",
                "исправен",
            ]
        )

    return di_rows, do_rows, ai_rows


def _col_letter(index: int) -> str:
    result = ""
    i = index
    while i > 0:
        i, rem = divmod(i - 1, 26)
        result = chr(65 + rem) + result
    return result


def _cell_xml(ref: str, value: object) -> str:
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f'<c r="{ref}"><v>{value}</v></c>'
    escaped = html.escape(str(value))
    return f'<c r="{ref}" t="inlineStr"><is><t>{escaped}</t></is></c>'


def _sheet_xml(rows: Iterable[list[object]]) -> str:
    row_xml: list[str] = []
    for r_index, row in enumerate(rows, start=1):
        cells = [
            _cell_xml(f"{_col_letter(c_index)}{r_index}", value)
            for c_index, value in enumerate(row, start=1)
        ]
        row_xml.append(f'<row r="{r_index}">{"".join(cells)}</row>')

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<sheetData>{"".join(row_xml)}</sheetData>'
        "</worksheet>"
    )


def render_protocol_excel(
    config: PlcConfig,
    ai_points: list[AiPoint],
    di_points: list[DiPoint],
    do_points: list[DoPoint],
    rtd_points: list[RtdPoint],
    output_path: str,
) -> None:
    di_rows, do_rows, ai_rows = _collect_rows(config, ai_points, di_points, do_points, rtd_points)
    created = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")

    content_types = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">
  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>
  <Default Extension=\"xml\" ContentType=\"application/xml\"/>
  <Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>
  <Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>
  <Override PartName=\"/xl/worksheets/sheet2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>
  <Override PartName=\"/xl/worksheets/sheet3.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>
  <Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>
  <Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>
</Types>
"""

    root_rels = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>
  <Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>
  <Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>
</Relationships>
"""

    workbook = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">
  <sheets>
    <sheet name=\"DI_Prot\" sheetId=\"1\" r:id=\"rId1\"/>
    <sheet name=\"DO_Prot\" sheetId=\"2\" r:id=\"rId2\"/>
    <sheet name=\"AI_Prot\" sheetId=\"3\" r:id=\"rId3\"/>
  </sheets>
</workbook>
"""

    workbook_rels = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>
  <Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet2.xml\"/>
  <Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet3.xml\"/>
</Relationships>
"""

    core = f"""<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">
  <dc:title>protocol</dc:title>
  <dc:creator>Generator1</dc:creator>
  <cp:lastModifiedBy>Generator1</cp:lastModifiedBy>
  <dcterms:created xsi:type=\"dcterms:W3CDTF\">{created}</dcterms:created>
  <dcterms:modified xsi:type=\"dcterms:W3CDTF\">{created}</dcterms:modified>
</cp:coreProperties>
"""

    app = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">
  <Application>Generator1</Application>
</Properties>
"""

    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", root_rels)
        zf.writestr("docProps/core.xml", core)
        zf.writestr("docProps/app.xml", app)
        zf.writestr("xl/workbook.xml", workbook)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        zf.writestr("xl/worksheets/sheet1.xml", _sheet_xml(di_rows))
        zf.writestr("xl/worksheets/sheet2.xml", _sheet_xml(do_rows))
        zf.writestr("xl/worksheets/sheet3.xml", _sheet_xml(ai_rows))
