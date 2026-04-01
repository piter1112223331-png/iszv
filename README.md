# MVP-1: Парсер XLSX-извещений об изменении

Минимально рабочая версия структурного разбора XLSX-извещений (без OCR, без LLM, без внешних API).

## Запуск

```bash
python -m parser.main "path/to/file.xlsx"
python -m parser.main "path/to/file.xlsx" --output result.json
python -m parser.main "path/to/file.xlsx" --debug --pretty
```

## Что доработано в текущем шаге

### 1) `notice_number`
- Извлечение делается только для `full` sheet (верхняя шапка).
- Кандидаты, похожие на `doc_code` (`ЕСРТ.0016.716.04121`), отбрасываются.
- Если уверенного номера нет — возвращается `null`.
- В debug выводятся `notice_candidates` и `rejected_notice_candidates`.

### 2) continuation sheets
- Разбор изменений теперь опирается на локальную шапку таблицы (`Изм.`; `Содержание изменения` может быть сокращено/отсутствовать).
- Если на continuation листе найдена таблица изменений, блоки извлекаются так же, как на full.

### 3) границы body
- Блок закрывается перед следующей валидной meta-signature (index + doc_code).
- Блок принудительно обрезается по stop-markers:
  - `Применяемость`, `Разослать`, `Составил`, `Проверил`, `Т. контроль`, `Н. контроль`, `Утвердил`, `Предст. заказ.`, `Изменения внес`, `Контрольную копию исправил`.
- Stop sections не попадают в `change_text`.

### 4) очистка `change_text`
- Схлопываются соседние дубли из merged cells.
- Удаляются одиночные мусорные `-`/`--`.
- Соседние повторяющиеся строки внутри блока убираются.

## Debug (`--debug`)
Для каждого листа выводится:
- `table_found`
- `table_header_row_start`
- `potential_meta_rows`
- `rejected_meta_rows`
- `reject_reasons`
- `stop_markers_hit`
- `blocks_closed_by_next_meta`
- `blocks_closed_by_stop_marker`
- `notice_candidates`
- `rejected_notice_candidates`
- `first_block_detected`
- `first_block_body_rows`
- `first_block_body_nonempty_cells`
- `first_block_closed_reason`
- `first_block_inline_text_before_cleanup`
- `first_block_inline_text_after_cleanup`
- `first_block_candidate_body_lines`
- `first_block_filtered_body_lines`
- `first_block_fallback_used`
- `first_block_final_text`

## Тесты

```bash
pytest -q
```

Тесты synthetic и не требуют реальных пользовательских XLSX в репозитории.
