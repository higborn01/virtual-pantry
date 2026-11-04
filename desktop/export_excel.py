"""'Export to Excel': rewrites the Pantry sheet rows of Meal_Plan_and_History.xlsx from the app's items.

Only columns A-I of the Pantry sheet are touched. Other sheets, the 'How to use' notes,
the Status formula, colours and drop-downs are kept. A backup copy is made first.
"""
import datetime as dt   # turn 'YYYY-MM-DD' text back into real dates for Excel
import os               # file paths
import shutil           # copying files (the backup)
from copy import copy   # copy a cell's look (style) onto another cell

from openpyxl import load_workbook                                     # read/write .xlsx
from openpyxl.formatting.formatting import ConditionalFormattingList  # Excel's colour rules

import db  # for the folder location

XLSX = os.path.join(os.path.dirname(db.HERE), "Meal_Plan_and_History.xlsx")            # the spreadsheet
BACKUP = os.path.join(os.path.dirname(db.HERE), "Meal_Plan_and_History.backup.xlsx")   # its backup copy
SHEET, TABLE = "Pantry", "Pantry"   # sheet name, and the name of the Excel table on it


def status_formula(r):
    """The spreadsheet's own Status formula for row r (so Excel keeps computing it)."""
    return f'=IF(C{r}="","?",IF(C{r}<=0,"Out",IF(C{r}<E{r},"Low","OK")))'


def export(items, path=XLSX, backup=BACKUP):
    """Write `items` to the Pantry sheet. Returns how many rows were written.
    Raises PermissionError if the file is open in Excel (the caller shows a friendly message)."""
    wb = load_workbook(path)                         # open the whole workbook (formulas kept as formulas)
    ws = wb[SHEET]                                   # the Pantry sheet
    table = ws.tables[TABLE]                         # the Excel table covering the item rows
    old_last = int(table.ref.split(":")[1][1:])      # last row of the table now, e.g. 'A1:I145' -> 145
    rows = sorted(items, key=lambda i: i["id"])      # original spreadsheet order (ids were given in that order); new items last
    last = len(rows) + 1                             # last row after export (row 1 is the header)

    template = [copy(ws.cell(2, c)._style) for c in range(1, 10)]  # look of row 2, columns A-I (yellow input cells etc.)
    blank = [copy(ws.cell(max(old_last, last) + 50, c)._style) for c in range(1, 10)]  # look of an empty cell far below

    for r, item in enumerate(rows, start=2):         # one row per item, starting under the header
        updated = dt.date.fromisoformat(item["updated"]) if item["updated"] else None  # text -> date (or blank)
        values = [item["name"], item["category"], item["qty"], item["unit"] or None, item["min_qty"],  # A-E
                  status_formula(r), item["location"], updated, item["notes"] or None]               # F-I
        for c, value in enumerate(values, start=1):  # each column A..I
            cell = ws.cell(r, c)                     # the cell at row r, column c
            if r > old_last:                         # a brand-new row: give it the same look as row 2
                cell._style = copy(template[c - 1])
            cell.value = value                       # write the value

    for r in range(last + 1, old_last + 1):          # rows left over if items were removed
        for c in range(1, 10):                       # columns A-I
            ws.cell(r, c).value = None               # empty the cell
            ws.cell(r, c)._style = copy(blank[c - 1])  # and make it look blank again

    table.ref = f"A1:I{last}"                        # resize the Excel table to the new row count
    if table.autoFilter is not None:                 # the table's filter buttons cover the same range
        table.autoFilter.ref = table.ref

    rules = ConditionalFormattingList()              # rebuild the colour rules with the new range
    for cf in ws.conditional_formatting:             # each existing rule group
        ref = str(cf.sqref)                          # the cells it covers, e.g. 'F2:F145'
        if ref.startswith("F2:"):                    # the Status column rule -> stretch to the new last row
            ref = f"F2:F{last}"
        for rule in cf.rules:                        # copy every rule in the group
            rules.add(ref, rule)
    ws.conditional_formatting = rules                # swap in the rebuilt rules

    shutil.copy2(path, backup)                       # back up the current file just before overwriting it
    wb.save(path)                                    # write the spreadsheet (PermissionError if Excel has it open)
    return len(rows)                                 # how many items were exported
