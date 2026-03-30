from __future__ import annotations

import html
import re
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable

from .xlsx_plc_reader import AiPoint, DiPoint, DoPoint, PlcConfig, ProtocolParams, RtdPoint


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

COLUMN_WIDTHS = [3.71, 39.57, 9.71, 12.71, 9.14, 12.71, 12.71, 16.14, 16.14, 12.57]


def _phase_letter(current_no: int) -> str:
    return {1: "A", 2: "B", 3: "C"}.get(current_no, "")


def _module_address(config: PlcConfig, io_type: str, module: int) -> str:
    if io_type == "ai":
        module_idx = module
    elif io_type == "rtd":
        module_idx = config.count_ai + module
    elif io_type == "di":
        module_idx = config.count_ai + config.count_rtd + module
    elif io_type == "do":
        module_idx = config.count_ai + config.count_rtd + config.count_di + module
    else:
        raise ValueError(f"Unsupported io_type: {io_type}")

    return f"A1_{module_idx + 1}"


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


def _di_contact(tag: str, protocol: ProtocolParams) -> int | str:
    normalized = tag.strip().lower()
    if re.fullmatch(r"qfds\d+", normalized) or re.fullmatch(r"qfd\d+", normalized):
        return protocol.num_qfd if protocol.num_qfd is not None else ""
    if re.fullmatch(r"qf\d+", normalized):
        return protocol.num_qf if protocol.num_qf is not None else ""
    if re.fullmatch(r"fd\d+", normalized):
        return protocol.num_fd if protocol.num_fd is not None else ""
    if re.fullmatch(r"km\d+(?:\.\d+)?", normalized):
        return 3
    if normalized in {"manu", "manuon"}:
        return protocol.num_cont_manu if protocol.num_cont_manu is not None else ""
    if normalized == "auto":
        return protocol.num_cont_auto if protocol.num_cont_auto is not None else ""
    return ""


def _collect_rows(
    config: PlcConfig,
    protocol: ProtocolParams,
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
                _module_address(config, "ai", point.module),
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
                _module_address(config, "rtd", point.module),
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
                _module_address(config, "di", point.module),
                point.channel,
                point.tag.upper(),
                _di_contact(point.tag, protocol),
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
                _module_address(config, "do", point.module),
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
    style = ' s="1"'
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f'<c r="{ref}"{style}><v>{value}</v></c>'
    escaped = html.escape(str(value))
    return f'<c r="{ref}" t="inlineStr"{style}><is><t>{escaped}</t></is></c>'


def _cols_xml(max_col: int) -> str:
    parts: list[str] = []
    for idx in range(1, max_col + 1):
        width = COLUMN_WIDTHS[idx - 1] if idx - 1 < len(COLUMN_WIDTHS) else 12.57
        parts.append(f'<col min="{idx}" max="{idx}" width="{width}" customWidth="1"/>')
    return "".join(parts)


def _sheet_xml(rows: Iterable[list[object]]) -> str:
    row_xml: list[str] = []
    max_col = 1
    for r_index, row in enumerate(rows, start=1):
        max_col = max(max_col, len(row))
        cells = [
            _cell_xml(f"{_col_letter(c_index)}{r_index}", value)
            for c_index, value in enumerate(row, start=1)
        ]
        row_xml.append(f'<row r="{r_index}">{"".join(cells)}</row>')

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<cols>{_cols_xml(max_col)}</cols>'
        f'<sheetData>{"".join(row_xml)}</sheetData>'
        "</worksheet>"
    )


def render_protocol_excel(
    config: PlcConfig,
    protocol: ProtocolParams,
    ai_points: list[AiPoint],
    di_points: list[DiPoint],
    do_points: list[DoPoint],
    rtd_points: list[RtdPoint],
    output_path: str,
) -> None:
    di_rows, do_rows, ai_rows = _collect_rows(config, protocol, ai_points, di_points, do_points, rtd_points)
    created = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")

    content_types = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">
  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>
  <Default Extension=\"xml\" ContentType=\"application/xml\"/>
  <Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>
  <Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>
  <Override PartName=\"/xl/worksheets/sheet2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>
  <Override PartName=\"/xl/worksheets/sheet3.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>
  <Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>
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
  <Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>
</Relationships>
"""

    styles = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">
  <fonts count=\"1\">
    <font><sz val=\"11\"/><name val=\"Times New Roman\"/></font>
  </fonts>
  <fills count=\"2\">
    <fill><patternFill patternType=\"none\"/></fill>
    <fill><patternFill patternType=\"gray125\"/></fill>
  </fills>
  <borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>
  <cellXfs count=\"2\">
    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>
    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\" wrapText=\"1\"/></xf>
  </cellXfs>
  <cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>
</styleSheet>
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
        zf.writestr("xl/styles.xml", styles)
        zf.writestr("xl/worksheets/sheet1.xml", _sheet_xml(di_rows))
        zf.writestr("xl/worksheets/sheet2.xml", _sheet_xml(do_rows))
        zf.writestr("xl/worksheets/sheet3.xml", _sheet_xml(ai_rows))
