# Generator1

Модульная заготовка генератора ST-кода из Excel.

## Что реализовано
- Модуль чтения листа `PLC` из `input_softitek.xlsx`.
- Модуль чтения листов `AI`, `RTD`, `DI`, `DO` и `Phase` (токи линий, состояния линий, RTD-точки и фазировка линий).
- Модуль генерации файла `GVL.txt` на основе параметров из `PLC`.
- Модуль генерации файла `Proc_IO.txt` для блока инициализации/обработки RTD, генерации токов линий (`AI/DO`), считывания состояний линий (`DI/DO`), записи команд на `DO` и инициализации фазировки.

## Структура
- `st_codegen/xlsx_plc_reader.py` — парсинг XLSX (без внешних библиотек), извлечение параметров PLC, AI/RTD/DI/DO-точек и фазировки.
- `st_codegen/gvl_generator.py` — рендер шаблона `GVL.txt`.
- `st_codegen/proc_io_generator.py` — рендер шаблона `Proc_IO.txt`.
- `generate_gvl.py` — CLI-скрипт для генерации `GVL.txt`.
- `generate_proc_io.py` — CLI-скрипт для генерации `Proc_IO.txt`.
- `generate_protocol.py` — CLI-скрипт для генерации `protocol.xlsx` (листы `DI_Prot`, `DO_Prot`, `AI_Prot`).

## Запуск
```bash
python generate_gvl.py input_softitek.xlsx GVL.txt
python generate_proc_io.py input_softitek.xlsx Proc_IO.txt
python generate_protocol.py input_softitek.xlsx protocol.xlsx
```

Если аргументы не передавать:
```bash
python generate_gvl.py
python generate_proc_io.py
python generate_protocol.py
```
по умолчанию используются `input_softitek.xlsx` и соответствующие выходные файлы.
