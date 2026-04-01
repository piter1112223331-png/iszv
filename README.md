# MVP-1: Парсер XLSX-извещений об изменении

Минимально рабочая версия структурного разбора XLSX-извещений (без OCR, без LLM, без внешних API).

## Структура

- `parser/workbook_loader.py` — загрузка книги XLSX.
- `parser/merged_cells.py` — чтение ячеек с учётом merged ranges.
- `parser/sheet_classifier.py` — отбор листов-кандидатов и классификация (`full` / `continuation` / `unknown`).
- `parser/sheet_parser.py` — парсинг одного листа (локальная шапка, номер листа, изменения).
- `parser/change_extractor.py` — извлечение блоков из области таблицы изменений.
- `parser/models.py` — dataclass-модели и сериализация в JSON.
- `parser/normalizer.py` — нормализация текста.
- `parser/main.py` — CLI точка входа.
- `tests/` — synthetic tests без реальных XLSX пользователя.

## Установка

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Запуск

```bash
python -m parser.main "path/to/file.xlsx"
python -m parser.main "path/to/file.xlsx" --output result.json
python -m parser.main "path/to/file.xlsx" --debug --pretty
```

## Что изменено по реальным прогонам

### 1) Защита `notice_number`

- Извлечение номера извещения теперь использует более строгий паттерн `Извещение ... № <код>`.
- Значения вроде `"Извещение"` / `"об изменении"` считаются невалидными.
- Номер должен содержать цифры (чтобы label не попадал в `notice_number`).

### 2) Явный поиск таблицы изменений

- Сначала ищется заголовочная зона таблицы по якорям:
  - `Изм.`
  - `Содержание изменения`
- Разбор `changes` идёт только ниже найденного заголовка таблицы.
- Строки выше таблицы не считаются meta-строками блоков.

### 3) Удаление дублей merged text

- При сборке текста строки схлопываются соседние одинаковые фрагменты.
- При сборке `change_text` повторяющиеся соседние строки также схлопываются.
- Это уменьшает дубли из merged cells.

## Debug (`--debug`)

Для каждого листа выводится:

- `sheet`;
- `sheet_kind`;
- `notice_number`;
- `sheet_no_detected`;
- `changes`;
- `table_header_row_start`;
- `potential_meta_rows`;
- `rejected_meta_rows`;
- `reject_reasons`.

## Что смотреть в JSON при первичной проверке

- `notice_number` (верхний уровень);
- `sheet_count_detected`;
- `sheets[].sheet_kind`;
- `sheets[].sheet_no_detected`;
- `sheets[].changes[].change_index`, `doc_code`, `change_text`;
- `all_changes`;
- `validation.errors` / `validation.warnings`.

## Тесты

```bash
pytest -q
```

Тесты синтетические и не требуют наличия реальных пользовательских XLSX в репозитории.
