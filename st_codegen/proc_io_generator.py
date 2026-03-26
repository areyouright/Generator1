from __future__ import annotations

from .xlsx_plc_reader import PhasePoint, PlcConfig, RtdPoint


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


def render_proc_io(config: PlcConfig, rtd_points: list[RtdPoint], phase_points: list[PhasePoint]) -> str:
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
