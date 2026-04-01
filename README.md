# MVP Parser: XLSX извещения об изменении

## Запуск

```bash
python -m parser.main "path/to/file.xlsx"
python -m parser.main "path/to/file.xlsx" --output result.json
python -m parser.main "path/to/file.xlsx" --debug --pretty
```

## MVP-3: слои результата

Полный JSON (совместим с MVP-1/2) сохраняется и дополняется.

Добавлены слои:
- `status`: `ok|warning|error`
- `summary`
- `llm_payload` (компактный слой для локальной LLM)
- `flowis_payload` (плоский слой для процесса)
- `validation_details` (машиночитаемые детали ошибок/предупреждений)

## LLM payload

Содержит:
- source_file
- notice_number
- sheet_count_detected
- document_header_compact (только непустые)
- changes (плоский список)
- warnings/errors
- summary

В `changes` оставлены только нужные поля + тех. флаги:
- sheet_no
- change_index
- doc_code
- change_text
- change_seq_global
- has_doc_code
- has_change_text
- text_length

## Flowis payload

Плоская структура:
- source_file
- status
- notice_number
- sheet_count_detected
- sheet_total_declared
- reason
- stock_instruction
- changes_count
- warnings
- errors

## Debug (`--debug`)

Показывает:
- header extraction debug
- validation warnings/errors
- computed_status
- summary
- llm_payload_preview
- flowis_payload_preview

## Тесты

```bash
pytest -q
```

Synthetic tests покрывают extraction, header extraction, validation и payload-слои MVP-3.
