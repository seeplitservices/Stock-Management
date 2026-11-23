import os
import gc
import logging
import openpyxl
from typing import Dict, List, Any

logger = logging.getLogger(__name__)


class UniversalSpreadsheetReader:
    """Reads raw 2D cell grids from Excel files safely and memory-efficiently for cloud hosting."""

    def iter_excel_sheets(self, file_path: str, max_rows_per_sheet: int = 15000):
        """Sequential generator that processes worksheets one at a time using OpenPyXL in read-only mode.
        Yields (sheet_name, sheet_idx, values) for each worksheet and immediately releases sheet memory.
        Explicitly closes the workbook handle in finally.
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Excel file not found at: {file_path}")

        lower_path = file_path.lower()
        if lower_path.endswith('.xls') and not lower_path.endswith('.xlsx'):
            raise ValueError(
                "Legacy .xls binary format is not supported directly. Please save or export your file as an Excel Workbook (.xlsx) before uploading."
            )

        wb = None
        try:
            # read_only=True and data_only=True streams XML and minimizes memory footprint
            wb = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
            sheet_names = list(wb.sheetnames)

            for sheet_idx, sheet_name in enumerate(sheet_names):
                ws = wb[sheet_name]
                values: List[List[Any]] = []
                consecutive_empty = 0

                for row in ws.iter_rows(values_only=True):
                    # Check if the row contains any non-empty cell
                    has_content = any(cell is not None and str(cell).strip() != "" for cell in row)
                    if has_content:
                        # Clean cells and trim trailing empty cells in this row to conserve memory
                        cleaned_row = [cell if cell is not None else "" for cell in row]
                        while cleaned_row and cleaned_row[-1] == "":
                            cleaned_row.pop()

                        values.append(cleaned_row)
                        consecutive_empty = 0
                    else:
                        consecutive_empty += 1
                        # If we have already started reading data and encounter 30 consecutive empty rows,
                        # break to avoid scanning 1,000,000 blank formatted rows in Excel.
                        if consecutive_empty >= 30 and len(values) > 0:
                            break

                    if len(values) >= max_rows_per_sheet:
                        logger.warning(f"Sheet '{sheet_name}' reached safety limit of {max_rows_per_sheet} rows.")
                        break

                yield sheet_name, sheet_idx, values

                # Release row memory for this worksheet immediately before proceeding to next worksheet
                del values
                del ws
                gc.collect()

        finally:
            if wb is not None:
                try:
                    wb.close()
                except Exception:
                    pass
                del wb
            gc.collect()

    def extract_from_excel(self, file_path: str, max_rows_per_sheet: int = 15000) -> Dict[str, Any]:
        """Extracts sheets dict from Excel file sequentially."""
        sheets_dict: Dict[str, Any] = {}
        for sheet_name, sheet_idx, values in self.iter_excel_sheets(file_path, max_rows_per_sheet):
            sheets_dict[sheet_name] = {
                "values": values,
                "row_count": len(values)
            }

        return {
            "source_type": "excel",
            "file_name": os.path.basename(file_path),
            "sheets": sheets_dict
        }
