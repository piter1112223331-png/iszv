# MVP-1: Парсер XLSX-извещений об изменении

Минимально рабочая версия структурного разбора XLSX-извещений (без OCR, без LLM, без внешних API).

## Структура

- `parser/workbook_loader.py` — загрузка книги XLSX.
- `parser/merged_cells.py` — чтение ячеек с учётом merged ranges.
- `parser/sheet_classifier.py` — отбор листов-кандидатов и классификация (`full` / `continuation` / `unknown`).
- `parser/sheet_parser.py` — парсинг одного листа (локальная шапка, номер листа, изменения).
- `parser/change_extractor.py` — эвристика повторяющихся блоков изменений.
- `parser/models.py` — dataclass-модели и сериализация в JSON.
- `parser/normalizer.py` — нормализация текста.
- `parser/main.py` — CLI точка входа.
- `tests/` — lightweight unit tests без реальных XLSX пользователя.

## Установка

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Запуск

Базовый запуск:

```bash
python -m parser.main "path/to/file.xlsx"
```

Сохранение JSON в файл:

```bash
python -m parser.main "path/to/file.xlsx" --output result.json
```

Режим локальной ручной диагностики:

```bash
python -m parser.main "path/to/file.xlsx" --debug --pretty
```

## Что показывает `--debug`

По каждому листу выводится (в `stderr`):

- имя листа;
- `sheet_kind`;
- найденный `notice_number`;
- найденный `sheet_no_detected`;
- число найденных `change blocks`;
- признак, что лист определён как кандидат.

После прохода выводится итог по документу:

- число candidate-листов;
- общее число `all_changes`.

## Что смотреть в JSON при первичной проверке

- `notice_number` (верхний уровень);
- `sheet_count_detected`;
- `sheets[].sheet_kind`;
- `sheets[].sheet_no_detected`;
- `sheets[].sheet_local_header.notice_number`;
- `sheets[].changes[].change_index`, `doc_code`, `change_text`, `zone_ref`;
- `all_changes`;
- `validation.errors` / `validation.warnings`.

## Что делает MVP-1

1. Читает все листы книги XLSX.
2. Отбирает листы-кандидаты по базовым маркерам:
   - `Извещение`
   - `Изм.`
   - `Содержание изменения`
3. Классифицирует лист:
   - `full` — если есть расширенные маркеры шапки (минимум 2 из набора)
   - `continuation` — если есть `Лист`, но нет расширенной шапки
   - `unknown` — если кандидат, но тип не распознан однозначно
4. Ищет повторяющиеся блоки изменений по более строгому признаку meta-строки:
   - наличие `Изм` + индекс изменения и/или код документа.
5. При сборке `change_text` не подтягивает строки заголовков таблицы.
6. Строит итоговый JSON документа и плоский `all_changes`.

## Ограничения MVP-1

- Эвристики текстовые (regex + ключевые слова), без глубокой бизнес-валидации.
- `document_header` пока заполняется `null`-полями.
- Распознавание `doc_code` и границ body-блока приблизительное и будет уточняться в MVP-2/3.

## Тесты

```bash
pytest -q
```

Тесты синтетические и не требуют наличия реальных пользовательских XLSX в репозитории.
