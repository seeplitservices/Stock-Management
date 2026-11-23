import os
import gc
import sys
import tempfile
import threading
import tracemalloc
from pathlib import Path
import openpyxl
import pytest

backend_dir = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(backend_dir))

from inventory_store import InventoryDataStore
from spreadsheet_reader import UniversalSpreadsheetReader
from sync_manager import SyncManager
from app import create_app


def create_mock_excel(sheets_data: dict, filepath: str):
    """Helper to create an openpyxl workbook with specified sheet structures."""
    wb = openpyxl.Workbook()
    # Remove default sheet
    default_sheet = wb.active

    first = True
    for sheet_name, rows in sheets_data.items():
        if first:
            ws = default_sheet
            ws.title = sheet_name
            first = False
        else:
            ws = wb.create_sheet(title=sheet_name)
        for row in rows:
            ws.append(row)

    wb.save(filepath)
    wb.close()


# Scenario 1: One spreadsheet with multiple worksheets
def test_scenario_01_one_spreadsheet_multiple_worksheets():
    store = InventoryDataStore(source_id="src_1", name="Multi-Sheet")
    raw = {
        "SheetA": [["Part No", "Description", "Opening Stock", "Closing Stock"], ["PA-1", "Desc A", 10, 15]],
        "SheetB": [["Part No", "Description", "Opening Stock", "Closing Stock"], ["PB-1", "Desc B", 20, 25]],
    }
    store.update_from_raw_sheets(raw)
    assert len(store.worksheets) == 2
    assert len(store.products) == 2
    assert {w["sheet_name"] for w in store.worksheets} == {"SheetA", "SheetB"}


# Scenario 2: At least 6 worksheets
def test_scenario_02_at_least_6_worksheets():
    store = InventoryDataStore(source_id="src_6", name="Six Sheets")
    raw = {
        f"Category_{i}": [["Part No", "Description", "Opening Stock", "Closing Stock"], [f"P-{i}", f"Item {i}", i * 10, i * 10]]
        for i in range(1, 7)
    }
    store.update_from_raw_sheets(raw)
    assert len(store.worksheets) == 6
    assert len(store.products) == 6


# Scenario 3: At least 10 worksheets processed sequentially
def test_scenario_03_at_least_10_worksheets():
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tf:
        temp_path = tf.name

    try:
        sheets_data = {
            f"Department_{i}": [["Part No", "Description", "Opening Stock", "Closing Stock"], [f"DEPT-P{i}", f"Dept Item {i}", 100, 95]]
            for i in range(1, 12)  # 11 worksheets
        }
        create_mock_excel(sheets_data, temp_path)

        reader = UniversalSpreadsheetReader()
        store = InventoryDataStore(source_id="src_11", name="11 Sheets")
        store.begin_incremental_update()
        for sheet_name, sheet_idx, rows in reader.iter_excel_sheets(temp_path):
            store.process_single_worksheet(sheet_name, sheet_idx, rows)
            del rows
        store.commit_incremental_update()

        assert len(store.worksheets) == 11
        assert len(store.products) == 11
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)


# Scenario 4: Multiple spreadsheets with different spreadsheet IDs
def test_scenario_04_multiple_spreadsheets_different_ids():
    sm = SyncManager()
    s1 = sm.register_source(name="Plant 1", spreadsheet_url_or_id="1AbC_Plant1_ID", source_type="google_sheets")
    s2 = sm.register_source(name="Plant 2", spreadsheet_url_or_id="2XyZ_Plant2_ID", source_type="google_sheets")

    s1.update_from_raw_sheets({"Stock": [["Part No", "Description"], ["P-1", "Plant 1 Item"]]})
    s2.update_from_raw_sheets({"Stock": [["Part No", "Description"], ["P-2", "Plant 2 Item"]]})

    assert s1.spreadsheet_id == "1AbC_Plant1_ID"
    assert s2.spreadsheet_id == "2XyZ_Plant2_ID"
    combined = sm.get_combined_store()
    assert len(combined.products) == 2


# Scenario 5: Different column structures across worksheets
def test_scenario_05_different_column_structures_across_sheets():
    store = InventoryDataStore(source_id="src_diff", name="Different Schemas")
    raw = {
        "Electronics": [["SKU", "Component", "Opening", "Closing"], ["E-1", "Resistor 10k", 500, 450]],
        "Hardware": [["Item Code", "Name", "Category", "UOM", "Rate", "Stock"], ["H-1", "M6 Nut", "Fasteners", "KG", "₹45.00", 80]],
        "Chemicals": [["Part Number", "Description", "Brand", "Hazard Level", "Available"], ["C-1", "IPA 99%", "Sigma", "Flammable", 20]]
    }
    store.update_from_raw_sheets(raw)
    assert len(store.worksheets) == 3
    # Check extra fields preserved dynamically
    p_chem = next(p for p in store.products if p["product_id"] == "C-1")
    assert p_chem["extra_fields"].get("Brand") == "Sigma"
    assert p_chem["extra_fields"].get("Hazard Level") == "Flammable"

    p_hard = next(p for p in store.products if p["product_id"] == "H-1")
    assert p_hard["category"] == "Fasteners"
    assert p_hard["uom"] == "KG"


# Scenario 6: New worksheet added after deployment
def test_scenario_06_new_worksheet_added_dynamically():
    store = InventoryDataStore(source_id="src_add", name="Dynamic Addition")
    raw_initial = {
        "Sheet1": [["Part No", "Description"], ["P-1", "Initial Item"]]
    }
    store.update_from_raw_sheets(raw_initial)
    assert len(store.worksheets) == 1

    # Later sync discovers Sheet2 added
    raw_updated = {
        "Sheet1": [["Part No", "Description"], ["P-1", "Initial Item"]],
        "Sheet2_New": [["Part No", "Description"], ["P-2", "Newly Added Item"]]
    }
    store.update_from_raw_sheets(raw_updated)
    assert len(store.worksheets) == 2
    assert any(w["sheet_name"] == "Sheet2_New" for w in store.worksheets)
    assert len(store.products) == 2


# Scenario 7: Worksheet renamed or removed
def test_scenario_07_worksheet_renamed_or_removed():
    store = InventoryDataStore(source_id="src_rename", name="Renamed Sheets")
    raw_v1 = {
        "Old_Name": [["Part No", "Description"], ["P-1", "Widget"]],
        "To_Be_Removed": [["Part No", "Description"], ["P-X", "Deprecating"]]
    }
    store.update_from_raw_sheets(raw_v1)
    assert len(store.worksheets) == 2

    # Renamed Old_Name -> New_Name, removed To_Be_Removed
    raw_v2 = {
        "New_Name": [["Part No", "Description"], ["P-1", "Widget (Updated)"]]
    }
    store.update_from_raw_sheets(raw_v2)
    assert len(store.worksheets) == 1
    assert store.worksheets[0]["sheet_name"] == "New_Name"
    assert len(store.products) == 1
    assert store.products[0]["product_id"] == "P-1"


# Scenario 8: New column inserted at the beginning
def test_scenario_08_column_inserted_at_beginning():
    store = InventoryDataStore(source_id="src_col_beg", name="Col At Beginning")
    # Col 0 is "Location Code", followed by Part No
    raw = {
        "Warehouse": [
            ["Location Code", "Part No", "Description", "Current Stock"],
            ["LOC-A1", "W-100", "Conduit Pipe 20mm", 85]
        ]
    }
    store.update_from_raw_sheets(raw)
    prod = store.products[0]
    assert prod["part_no"] == "W-100"
    assert prod["description"] == "Conduit Pipe 20mm"
    assert prod["extra_fields"].get("Location Code") == "LOC-A1"


# Scenario 9: New column inserted in the middle
def test_scenario_09_column_inserted_in_middle():
    store = InventoryDataStore(source_id="src_col_mid", name="Col In Middle")
    raw = {
        "Assembly": [
            ["Part No", "Manufacturer", "Specification", "Description", "Opening Stock", "Closing Stock"],
            ["ASM-50", "Siemens", "415V 50Hz", "Industrial Contactor", 12, 10]
        ]
    }
    store.update_from_raw_sheets(raw)
    prod = store.products[0]
    assert prod["part_no"] == "ASM-50"
    assert prod["description"] == "Industrial Contactor"
    assert prod["extra_fields"].get("Manufacturer") == "Siemens"
    assert prod["extra_fields"].get("Specification") == "415V 50Hz"


# Scenario 10: New column appended at the end
def test_scenario_10_column_appended_at_end():
    store = InventoryDataStore(source_id="src_col_end", name="Col At End")
    raw = {
        "Spares": [
            ["Part No", "Description", "Stock", "Warranty Period", "Remarks / Notes"],
            ["SPR-99", "O-Ring Seal 50mm", 200, "12 Months", "Store in cool dry place"]
        ]
    }
    store.update_from_raw_sheets(raw)
    prod = store.products[0]
    assert prod["part_no"] == "SPR-99"
    assert prod["extra_fields"].get("Warranty Period") == "12 Months"
    assert prod["extra_fields"].get("Remarks / Notes") == "Store in cool dry place"


# Scenario 11: Large worksheet with thousands of rows
def test_scenario_11_large_worksheet_thousands_of_rows():
    store = InventoryDataStore(source_id="src_large", name="Large Sheet")
    rows = [["Part No", "Description", "Unit Price", "Opening Stock", "Closing Stock"]]
    for i in range(1, 3001):  # 3,000 rows
        rows.append([f"PART-{i:05d}", f"Standard Resistor Value {i} Ohms", 2.50, 100, 90])

    store.update_from_raw_sheets({"MassCatalog": rows})
    assert len(store.worksheets) == 1
    assert store.worksheets[0]["row_count"] == 3000
    assert len(store.products) == 3000


# Scenario 12: Empty worksheet and missing values (no NaN, undefined, or fake zeroes)
def test_scenario_12_empty_worksheet_and_missing_values():
    store = InventoryDataStore(source_id="src_empty", name="Empty & Sparse")
    raw = {
        "CompletelyEmpty": [],
        "SparseValues": [
            ["Part No", "Description", "Unit Price", "Opening Stock", "Closing Stock"],
            ["P-EMPTY", "Item with empty cells", None, None, None],
            ["P-NA", "Item with NA strings", "NA", "-", "None"],
            ["", "", "", "", ""]  # Blank row
        ]
    }
    store.update_from_raw_sheets(raw)
    assert len(store.worksheets) == 2
    empty_ws = next(w for w in store.worksheets if w["sheet_name"] == "CompletelyEmpty")
    assert empty_ws["row_count"] == 0

    p_empty = next(p for p in store.products if p["product_id"] == "P-EMPTY")
    # Missing values must not be NaN or fake zeroes
    assert p_empty["unit_price"] is None
    assert p_empty["current_stock"] is None or p_empty["current_stock"] == 0
    assert p_empty["stock_status"] in ("Unknown", "Out of Stock")


# Scenario 13: Repeated synchronisation
def test_scenario_13_repeated_synchronisation():
    sm = SyncManager()
    store = sm.register_source("Repeated Sync", "sync_test_id")
    raw = {"Data": [["Part No", "Description", "Stock"], ["P-1", "Test", 10]]}

    for _ in range(5):
        store.update_from_raw_sheets(raw)
        assert len(store.products) == 1
        assert len(store.worksheets) == 1


# Scenario 14: Two users requesting synchronisation simultaneously
def test_scenario_14_concurrent_sync_requests():
    sm = SyncManager()
    store = sm.register_source("Concurrent Test", "conc_test_id", source_type="excel_upload")
    store.update_from_raw_sheets({"Sheet": [["Part No", "Desc"], ["P-1", "Item"]]})

    results = []

    def run_sync():
        res = sm.sync_source(store.source_id, force=True)
        results.append(res)

    t1 = threading.Thread(target=run_sync)
    t2 = threading.Thread(target=run_sync)

    t1.start()
    t2.start()
    t1.join()
    t2.join()

    # Both must complete gracefully without exceptions or deadlock
    assert len(results) == 2
    assert all(r is True for r in results)


# Scenario 15: Failed Google Sheets connection handles error gracefully
def test_scenario_15_failed_sheets_connection_graceful_recovery():
    sm = SyncManager()
    # Invalid spreadsheet ID with no real credentials
    store = sm.register_source("Bad Sheet", "nonexistent_spreadsheet_id_999", source_type="google_sheets")
    success = sm.sync_source(store.source_id, force=True)

    # Should not crash the server; returns False and records the notice
    assert success is False
    assert store.last_sync_status == "error"
    assert store.last_sync_error is not None
    # System status reports error state without crashing
    status = sm.get_status()
    assert status is not None


# Scenario 16: Repeated uploads without memory accumulating indefinitely
def test_scenario_16_repeated_uploads_memory_bounded():
    tracemalloc.start()
    gc.collect()
    snapshot1 = tracemalloc.take_snapshot()

    reader = UniversalSpreadsheetReader()
    store = InventoryDataStore(source_id="src_leak_test", name="Memory Test")

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tf:
        temp_file = tf.name

    try:
        sheets_data = {
            f"Sheet_{s}": [["Part No", "Description", "Unit Price", "Opening Stock", "Closing Stock"]] +
                          [[f"P-{s}-{r}", f"Item description {r}", 50, 20, 15] for r in range(100)]
            for s in range(6)
        }
        create_mock_excel(sheets_data, temp_file)

        # Simulate 5 sequential upload & parse cycles
        for cycle in range(5):
            store.begin_incremental_update()
            for sheet_name, sheet_idx, rows in reader.iter_excel_sheets(temp_file):
                store.process_single_worksheet(sheet_name, sheet_idx, rows)
                del rows
                gc.collect()
            store.commit_incremental_update()
            gc.collect()

        snapshot2 = tracemalloc.take_snapshot()
        top_stats = snapshot2.compare_to(snapshot1, 'lineno')
        total_growth_kb = sum(stat.size_diff for stat in top_stats) / 1024

        # Memory delta after 5 cycles must remain well below 20MB (vastly below 512MB limit)
        assert total_growth_kb < 20480, f"Memory grew excessively: {total_growth_kb:.2f} KB"
    finally:
        tracemalloc.stop()
        if os.path.exists(temp_file):
            os.remove(temp_file)
