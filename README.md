# MVP Parser: XLSX извещения об изменении

## Запуск

```bash
python -m parser.main "path/to/file.xlsx"
python -m parser.main "path/to/file.xlsx" --output result.json
python -m parser.main "path/to/file.xlsx" --debug --pretty
```

## Flowis integration (CLI)

По умолчанию отправляется `llm_payload` (не full JSON):

```bash
python -m parser.main "1.xlsx" --pretty --send-flowis --flowis-url "http://localhost:3000/api/v1/prediction/..."
python -m parser.main "1.xlsx" --send-flowis --flowis-payload flowis --flowis-save-response flowis_response.json
```

Опции:
- `--send-flowis` — отправить payload в Flowis после обычного парсинга
- `--flowis-url` — endpoint URL (если не указан, используется `FLOWIS_API_URL`)
- `--flowis-payload {llm,flowis,full}` — какой слой отправлять, default: `llm`
- `--flowis-timeout <sec>` — timeout запроса, default: `180`
- `--flowis-save-response <path>` — сохранить JSON-ответ Flowis в файл

Если `--flowis-url` и `FLOWIS_API_URL` отсутствуют, CLI завершится с понятной ошибкой.

## MVP-3+ business attributes

`document_header` расширен полями формы:
- developer
- notice_number
- reason
- code
- sheet_no_declared
- sheet_total_declared
- stock_instruction
- implementation_instruction
- applicability
- distribution
- release_center / release_date (если извлекаются)

Добавлен top-level `approvals`:
- author
- reviewer
- norm_control
- approver

## Layered payloads

- Full JSON (подробный слой) остаётся совместимым
- `llm_payload` (compact)
- `flowis_payload` (flat)
- `validation_details` (machine-readable)
- `status`, `summary`

## Debug (`--debug`)

Включает:
- header field diagnostics
- approvals_found / approvals_missing
- developer/notice/code/sheet_no candidates
- sheet_total_declared_candidate
- approvals_candidates / approvals_rejected_reasons
- computed_status
- summary
- llm_payload_preview
- flowis_payload_preview

## Тесты

```bash
pytest -q
```
