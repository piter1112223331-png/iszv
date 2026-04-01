from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook


class WorkbookLoadError(RuntimeError):
    pass


def load_xlsx(path: str | Path) -> Workbook:
    file_path = Path(path)
    if not file_path.exists():
        raise WorkbookLoadError(f"File not found: {file_path}")
    if file_path.suffix.lower() != ".xlsx":
        raise WorkbookLoadError("Only .xlsx files are supported in MVP-1")

    try:
        return load_workbook(filename=file_path, data_only=True)
    except Exception as exc:  # noqa: BLE001
        raise WorkbookLoadError(f"Failed to load workbook: {exc}") from exc
