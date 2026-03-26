# Generator1

Модульная заготовка генератора ST-кода из Excel.

## Что реализовано
- Модуль чтения листа `PLC` из `input_softitek.xlsx`.
- Модуль генерации файла `GVL.txt` на основе параметров из `PLC`.

## Структура
- `st_codegen/xlsx_plc_reader.py` — парсинг XLSX (без внешних библиотек) и извлечение параметров PLC.
- `st_codegen/gvl_generator.py` — рендер шаблона `GVL.txt`.
- `generate_gvl.py` — CLI-скрипт для запуска генерации.

## Запуск
```bash
python generate_gvl.py input_softitek.xlsx GVL.txt
```

Если аргументы не передавать:
```bash
python generate_gvl.py
```
по умолчанию используются `input_softitek.xlsx` и `GVL.txt`.
