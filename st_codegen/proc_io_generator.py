from __future__ import annotations

import re

from .xlsx_plc_reader import AiPoint, DoPoint, PhasePoint, PlcConfig, RtdPoint


def _module_address(config: PlcConfig, module_idx: int) -> str:
    return f"A1_{module_idx + 1}"


def _module_type_comment(config: PlcConfig, module_idx: int) -> str:
    if module_idx <= config.count_ai:
        return "AI-модуль"
    if module_idx <= config.count_ai + config.count_rtd:
        return "RTD-модуль"
    if module_idx <= config.count_ai + config.count_rtd + config.count_di:
        return "DI-модуль"
    return "DO-модуль"


def _module_address_for_rtd(config: PlcConfig, rtd_module: int) -> str:
    module_position = config.count_ai + rtd_module
    return _module_address(config, module_position)


def _parse_km_tag(tag: str) -> tuple[int, int]:
    normalized = tag.strip().lower()
    match = re.fullmatch(r"km(\d+)(?:\.(\d+))?", normalized)
    if not match:
        raise ValueError(f"Unsupported DO line tag: '{tag}'")
    km_no = int(match.group(1))
    variant = int(match.group(2)) if match.group(2) else 1
    return km_no, variant


def _build_ai_lookup(ai_points: list[AiPoint]) -> dict[tuple[int, int], AiPoint]:
    lookup: dict[tuple[int, int], AiPoint] = {}
    for point in ai_points:
        tag = point.tag.strip().lower()
        ta_match = re.fullmatch(r"ta(\d+)\.(\d+)", tag)
        if ta_match:
            line_no = int(ta_match.group(1))
            current_no = int(ta_match.group(2))
            if current_no in (1, 2, 3):
                lookup[(line_no, current_no)] = point
            continue

        tan_match = re.fullmatch(r"tan(\d+)", tag)
        if tan_match:
            line_no = int(tan_match.group(1))
            lookup[(line_no, 4)] = point
    return lookup


def _render_scaled_current(target_line_no: int, current_no: int, source: AiPoint) -> str:
    hi_limit = 50 if current_no == 4 else 63
    return (
        f"GVL.DataPack.Lines[{target_line_no}].Currents[{current_no}] := "
        f"OSCAT_BASIC.SCALE_R(X := WORD_TO_REAL(GVL.AIs[{source.module}].awValue[{source.channel}]), "
        f"I_LO := 6520, I_HI := 32768, O_LO := 0, O_HI := {hi_limit});"
    )


def _append_ai_currents(lines: list[str], config: PlcConfig, ai_points: list[AiPoint], do_points: list[DoPoint]) -> None:
    ai_lookup = _build_ai_lookup(ai_points)
    km_points = [point for point in do_points if re.fullmatch(r"km\d+(?:\.\d+)?", point.tag.strip().lower())]
    source_line_by_base: dict[int, int] = {}

    lines.append("// Инициализируем токи по линиям")
    line_no = 0
    for point in km_points:
        if line_no >= config.count_lines:
            break
        line_no += 1
        km_no, variant = _parse_km_tag(point.tag)
        lines.append(f"// {point.tag.strip().upper()}")

        if variant == 1:
            source_line_by_base[km_no] = line_no
            for current_no in (1, 2, 3, 4):
                source = ai_lookup.get((km_no, current_no))
                if source is None:
                    continue
                lines.append(_render_scaled_current(line_no, current_no, source))
        else:
            source_line = source_line_by_base.get(km_no)
            if source_line is None:
                for current_no in (1, 2, 3, 4):
                    source = ai_lookup.get((km_no, current_no))
                    if source is None:
                        continue
                    lines.append(_render_scaled_current(line_no, current_no, source))
                source_line_by_base[km_no] = line_no
            else:
                available_currents = [
                    current_no for current_no in (1, 2, 3, 4) if (km_no, current_no) in ai_lookup
                ]
                for current_no in (1, 2, 3, 4):
                    if available_currents and current_no not in available_currents:
                        continue
                    lines.append(
                        f"GVL.DataPack.Lines[{line_no}].Currents[{current_no}] := "
                        f"GVL.DataPack.Lines[{source_line}].Currents[{current_no}];"
                    )
        lines.append("")


def render_proc_io(
    config: PlcConfig,
    ai_points: list[AiPoint],
    do_points: list[DoPoint],
    rtd_points: list[RtdPoint],
    phase_points: list[PhasePoint],
) -> str:
    lines: list[str] = []
    lines.append("//------------------------- Первая инициализируемая программа, в которой считываются данные от модулей -------------------------")
    lines.append("(*----------------------------Расчёт ошибки датчика температуры -----------------------------------*)")
    lines.append("// Инициализация датчиков и проверка их состояния")
    lines.append("")

    for rtd_module in range(1, config.count_rtd + 1):
        module_addr = _module_address_for_rtd(config, rtd_module)
        lines.append(f"TOF_AlarmRTD[{rtd_module}](IN := {module_addr}.xError, PT := T#2S);")

    lines.append("")

    max_dt_index = config.count_dt
    selected_points = rtd_points[: max_dt_index + 1]

    for dt_index, point in enumerate(selected_points):
        lines.append(f"// Датчик {dt_index} ({point.tag})")
        lines.append(
            f"GVL.DataPack.aDT[{dt_index}].rT := GVL.RTDs[{point.module}].arValue[{point.channel}];"
        )
        lines.append(
            f"tonErr[{dt_index}](IN := GVL.RTDs[{point.module}].awStatus[{point.channel}] <> 0, PT := T#2S);"
        )
        lines.append(
            f"GVL.DataPack.aDT[{dt_index}].wErrDT := TO_WORD(tonErr[{dt_index}].Q OR TOF_AlarmRTD[{point.module}].Q);"
        )

    lines.append("")
    _append_ai_currents(lines, config, ai_points, do_points)

    lines.append("// Инициализируем фазировку")
    phase_map = {point.line_no: point.phase_code for point in phase_points}
    for line_no in range(1, config.count_lines + 1):
        phase_code = phase_map.get(line_no, 0)
        lines.append(f"GVL.inNumL[{line_no}] := {phase_code};")

    lines.append("")
    lines.append("// Инициализируем ошибки модулей")
    for module_idx in range(1, config.count_modules + 1):
        lines.append(
            f"GVL.Err_mod[{module_idx}]\t:=\t{_module_address(config, module_idx)}.xError;\t// {_module_type_comment(config, module_idx)}"
        )

    return "\n".join(lines) + "\n"
