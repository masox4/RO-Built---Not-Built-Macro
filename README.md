## RO Built / Not Built macro module

This repo includes `RO_StockList_Macros.bas` with button and import/filter automation.

### What it does

- `SetupActionButtons` (or `SetupClearButtons`) creates small buttons with white text:
  - `RO StockList!F1` -> **Clear Old Data**
  - `RO StockList!K1` -> **Import RO StockList**
  - `RO StockList!L1` -> **Show NOT Built RO Poducts** (dark green)
  - `RO StockList!M1` -> **Undo Show NOT Built**
  - `RAW RO StockList!E1` -> **Clear Old Data**
- `ClearOldData_Main` clears `RO StockList` from `A2:AD` until the first blank row in column `A`.
- `ClearOldData_Raw` clears `RAW RO StockList` from `A2:AE` until the first blank row in column `A`.
- `ImportROStockList`:
  - Requires `RO StockList` data to be cleared first.
  - Copies `RAW RO StockList` rows from `A2:AA` down to first blank SKU row.
  - Pastes to `RO StockList` with an inserted formula column `B` (Built/Not Built).
  - Adds formulas in `AC` (Ignore) and `AD` (Missed) down to the imported last row.
  - Uses paste of values + number formats to preserve numeric/text formatting (including EAN-style fields).
- `ShowNotBuiltROProducts`:
  - Filters `AC` to `-` only.
  - Filters `AD` to `-` only.
  - Filters `B` to exclude `Built`.
- `UndoShowNotBuilt`:
  - Shows all rows and removes filters.

### How to use in your `.xlsm`

1. Open `EYES vs RO What Needs Building.xlsm` in Excel.
2. Press `ALT+F11` to open the VBA editor.
3. Right-click the workbook project -> **Import File...** -> choose `RO_StockList_Macros.bas`.
4. Run `SetupActionButtons` once.

Then use the sheet buttons as needed.
