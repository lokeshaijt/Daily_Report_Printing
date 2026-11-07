from .processor import Options, Sheet, display_value, process
from .reader import read_workbook
from .writer import write_workbook


def fix_report(data: bytes, filename: str = "", options: Options | None = None):
    """Upload bytes -> (processed sheets, corrected .xlsx bytes)."""
    sheets = process(read_workbook(data, filename), options)
    return sheets, write_workbook(sheets)


__all__ = ["Options", "Sheet", "display_value", "process", "read_workbook", "write_workbook", "fix_report"]
