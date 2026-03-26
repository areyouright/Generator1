from __future__ import annotations

from .xlsx_plc_reader import PlcConfig


def render_gvl(config: PlcConfig) -> str:
    return f"""// Generated GVL
{{attribute 'qualified_only'}}
VAR_GLOBAL RETAIN
\tDataPack:\t\t\t\tstDataPack;\t\t\t// Массив всех данных для передачи по типу данных 4x
END_VAR

VAR_GLOBAL CONSTANT
\tgc_countL:\t\t\t\tUINT\t\t:= {config.count_lines};\t// Кол-во линий
\tgc_countDt:\t\t\t\tUINT\t\t:= {config.count_dt};\t// Количество датчиков т-ры, отсчитывая от нуля
\tgc_count_modules:\t\tUINT\t\t:= {config.count_modules};\t// Общее количество модулей
\tgc_first_group:\t\t\tUINT\t\t:= {config.first_group};\t// Кол-во линий в первой группе
\tgc_countRegs:\t\t\tUINT\t\t:= 400;\t// Кол-во регистров в DataPack'e - выставляем с небольшим запасом
\tcount_ai:\t\t\t\tUINT\t\t:= {config.count_ai};\t// Количество AI-модулей
\tcount_rtd:\t\t\t\tUINT\t\t:= {config.count_rtd};\t// Количество RTD-модулей
\tcount_di:\t\t\t\tUINT\t\t:= {config.count_di};\t// Количество DI-модулей
\tcount_do:\t\t\t\tUINT\t\t:= {config.count_do};\t// Количество DO-модулей
END_VAR

VAR_GLOBAL

\tinNumL:\t\t\t\t\tARRAY [1..gc_countL] OF UINT;\t\t\t\t//массив фазировки линий
\talarmQFD:\t\t\t\tARRAY [1..gc_countL] OF BOOL;\t\t\t\t// массив сработок дифф. автоматов
\tErr_mod:\t\t\t\tARRAY [1..gc_count_modules] OF BOOL;\t\t// массив ошибок модулей

(*----------------------------------------------------------*)
\tAIs \t:\tARRAY [1..count_ai]\t\tOF stAI;\t\t// AI-модули (Софтитек)
\tRTDs\t:\tARRAY [1..count_rtd]\tOF stRTD;\t\t// RTD-модули (Софтитек)
\tDIs\t\t:\tARRAY [1..count_di]\t\tOF WORD;\t\t// DI-модули (Софтитек)
\tDOs\t\t:\tARRAY [1..count_do]\t\tOF WORD;\t\t// DO-модули (Софтитек)
\t
END_VAR
"""
