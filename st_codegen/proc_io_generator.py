from __future__ import annotations

from .xlsx_plc_reader import PlcConfig, RtdPoint


def _module_address_for_rtd(config: PlcConfig, rtd_module: int) -> str:
    module_position = config.count_ai + rtd_module
    return f"A1_{module_position + 1}"


def render_proc_io(config: PlcConfig, rtd_points: list[RtdPoint]) -> str:
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

    return "\n".join(lines) + "\n"
