# MVP Parser: XLSX извещения об изменении

## Запуск

```bash
python -m parser.main "path/to/file.xlsx"
python -m parser.main "path/to/file.xlsx" --output result.json
python -m parser.main "path/to/file.xlsx" --debug --pretty
```

## MVP-2 (поверх MVP-1)

### 1) `document_header` extraction
Извлечение выполняется из **первого full-sheet** по anchor-полям:
- sender
- reason
- code
- sheet_total_declared
- release_center
- release_date
- stock_instruction
- implementation_instruction
- applicability
- distribution

Если надёжного значения нет — поле остаётся `null`.

### 2) Rule-based validation
Формируются:
- `validation.errors`
- `validation.warnings`

Ошибки:
- `no_candidate_sheets`
- `no_change_blocks`
- `empty_doc_code:<label>`
- `empty_change_index:<label>`

Warnings:
- `notice_number_missing`
- `header_field_missing:<field>`
- `empty_change_text:<label>`
- `sheet_total_declared_missing`
- `sheet_count_mismatch`
- `suspicious_block_boundary:<label>`

## Debug (`--debug`)
Кроме sheet/block диагностики MVP-1 выводится:
- `header_found`
- `header_missing`
- `validation_errors`
- `validation_warnings`

Для first block по-прежнему доступны:
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

Synthetic tests покрывают extraction, header extraction и validation.
