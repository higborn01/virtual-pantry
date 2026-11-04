"""One-time (re-runnable) import from Meal_Plan_and_History.xlsx into pantry.db.

Run:  python import_excel.py

- Pantry items: only NEW names are added; items already in the database are left
  alone, so changes made in the app are never overwritten.
- Dishes: added or refreshed from Meal History + Dinner Wheel.
- Dish ingredients: rebuilt from scratch every run (they aren't edited in the app).
"""
import datetime as dt  # date types, to recognise date cells
import os              # file-path tools
import re              # regular expressions (text pattern matching)
import sys             # sys.exit() to stop with a message

from openpyxl import load_workbook  # reads .xlsx files

import db  # our own db.py (connect, DB folder)

XLSX = os.path.join(os.path.dirname(db.HERE), "Meal_Plan_and_History.xlsx")  # spreadsheet one folder up

# Words stripped from the front of an ingredient to get its keyword.
UNITS = {  # a set: fast "is this word in here?" checks
    "can", "cans", "cup", "cups", "tbsp", "tsp", "lb", "lbs", "oz", "bag", "bags",
    "box", "boxes", "jar", "jars", "bottle", "bottles", "pkg", "package", "packages",
    "clove", "cloves", "bunch", "head", "heads", "pint", "pints", "dozen", "block",
    "blocks", "stick", "sticks", "slice", "slices", "of",
}
QTY = re.compile(r"^[\d½⅓⅔¼¾⅛.\-/]+$")  # a word made only of digits/fractions: '2', '1/2', '¾', '1-2'


# ------------------------------------------------------------ text helpers
def split_list(text):
    """Split a shopping list on commas, but not commas inside (parentheses)."""
    parts, depth, cur = [], 0, ""   # finished pieces, how deep in ( ) we are, piece being built
    for ch in text:                 # walk the text one character at a time
        if ch == "(":               # entering brackets
            depth += 1
        elif ch == ")":             # leaving brackets (never below 0)
            depth = max(0, depth - 1)
        if ch == "," and depth == 0:  # a comma outside brackets ends a piece
            parts.append(cur)         # save the piece
            cur = ""                  # start a new one
        else:
            cur += ch                 # any other character joins the current piece
    parts.append(cur)                 # save the last piece (no comma after it)
    return [p.strip() for p in parts if p.strip()]  # trim spaces, drop empty pieces


def singular(word):
    """Rough plural -> singular for the last word: beans->bean, tomatoes->tomato."""
    if word.endswith(("ss", "us")) or len(word) <= 3:  # 'swiss', 'hummus', 'peas'-short words: leave alone
        return word
    if word.endswith("ies"):    # 'berries' -> 'berry'
        return word[:-3] + "y"
    if word.endswith("oes"):    # 'tomatoes' -> 'tomato'
        return word[:-2]
    if word.endswith("s"):      # 'beans' -> 'bean'
        return word[:-1]
    return word                 # already singular


def keyword(text):
    """'1 can Black Beans' -> 'black bean'. Matching gets smarter in Step 7."""
    words = text.lower().split()  # lower-case, split into words
    while words and (QTY.match(words[0]) or words[0] in UNITS):  # while the first word is a number or unit...
        words.pop(0)                                             # ...remove it
    if not words:                 # nothing left (e.g. the text was just '2 cups')
        return ""
    words[-1] = singular(words[-1])  # make the last word singular
    return " ".join(words)           # glue the words back together


def parse_ingredients(text):
    """Yield (raw, keyword, alt_group) for each ingredient in a shopping list."""
    for group, raw in enumerate(split_list(text or "")):     # each comma piece, numbered 0, 1, 2...
        no_parens = re.sub(r"\([^)]*\)", " ", raw)           # drop bracketed text like '(or carton)'
        for alt in re.split(r"\s+or\s+|/", no_parens):        # split 'feta or mozzarella' / 'lemon/lime'
            kw = keyword(alt)                                 # normalize each alternative
            if kw:                                            # skip empties
                yield raw, kw, group                          # hand back one ingredient row


def clean(value):
    return str(value).strip() if value is not None else ""  # cell -> trimmed text ('' if blank)


def to_number(value):
    """Spreadsheet cell -> float, or None if blank / not a number."""
    if value is None or value == "":  # blank cell
        return None
    try:
        return float(value)           # 1 -> 1.0, '0.5' -> 0.5
    except (TypeError, ValueError):   # text that isn't a number
        return None


def to_date(value):
    if isinstance(value, (dt.datetime, dt.date)):  # Excel date cell
        return value.strftime("%Y-%m-%d")          # store as 'YYYY-MM-DD' text
    return None                                    # blank or not a date -> unknown


# ------------------------------------------------------------ importers
def import_items(conn, ws):
    """Columns A-I: Item, Category, Qty, Unit, Min, Status (skip), Location, Updated, Notes."""
    added = 0                                                            # count of new items
    for row in ws.iter_rows(min_row=2, max_col=9, values_only=True):    # every row after the header, cols A-I
        name = clean(row[0])                                             # column A: item name
        if not name:                                                     # skip blank rows
            continue
        # INSERT OR IGNORE: if the name already exists (UNIQUE), skip it silently.
        # The ? marks are placeholders: sqlite3 fills them safely from the tuple below.
        cur = conn.execute(
            """INSERT OR IGNORE INTO items
               (name, category, qty, unit, min_qty, location, updated, notes)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (name, clean(row[1]) or "Other", to_number(row[2]), clean(row[3]),   # A, B, C, D
             to_number(row[4]) or 0, clean(row[6]) or "Pantry", to_date(row[7]),  # E, G, H (F = Status, skipped)
             clean(row[8])),                                                      # I
        )
        added += cur.rowcount  # 1 if inserted, 0 if ignored
    return added               # how many were new


def collect_dishes(wb):
    """Merge both dish sheets into one dict keyed by lower-case name."""
    dishes = {}  # lower-case name -> dish dict
    for row in wb["Meal History"].iter_rows(min_row=2, max_col=6, values_only=True):  # cols A-F
        name = clean(row[0])           # column A: dish
        if not name:                   # skip blank rows
            continue
        status = clean(row[3])         # column D: status
        dishes[name.lower()] = dict(   # build the dish record
            name=name, category=clean(row[1]), cuisine=clean(row[2]), status=status,
            notes=clean(row[4]), wheel_category=None,
            made=int(status.startswith(("Made before", "Likely made"))),  # True/False -> 1/0
            source="Meal History", shopping_list=clean(row[5]),
        )
    for row in wb["Dinner Wheel"].iter_rows(min_row=2, max_col=4, values_only=True):  # cols A-D
        name = clean(row[0])                          # column A: dish
        if not name:                                  # skip blank rows
            continue
        made = int(clean(row[2]).lower() == "yes")   # column C: 'Yes' -> 1, else 0
        d = dishes.get(name.lower())                 # already seen in Meal History?
        if d:  # in both sheets: keep Meal History details, add the wheel info
            d.update(wheel_category=clean(row[1]), made=max(d["made"], made), source="Both")
            d["shopping_list"] = d["shopping_list"] or clean(row[3])  # use the wheel's list only if Meal History had none
        else:  # only on the Dinner Wheel: make a new record
            dishes[name.lower()] = dict(
                name=name, category="Dinner", cuisine="", status="Wheel", notes="",
                wheel_category=clean(row[1]), made=made, source="Dinner Wheel",
                shopping_list=clean(row[3]),
            )
    return list(dishes.values())  # just the dish records, as a list


def import_dishes(conn, dishes):
    for d in dishes:  # each merged dish
        # "Upsert": insert, or if the name exists, update that row instead.
        # :name etc. are named placeholders, filled from the dict d.
        # "excluded" means the values we just tried to insert.
        conn.execute(
            """INSERT INTO dishes (name, category, cuisine, status, notes,
                                   wheel_category, made, source, shopping_list)
               VALUES (:name, :category, :cuisine, :status, :notes,
                       :wheel_category, :made, :source, :shopping_list)
               ON CONFLICT(name) DO UPDATE SET
                 category=excluded.category, cuisine=excluded.cuisine,
                 status=excluded.status, notes=excluded.notes,
                 wheel_category=excluded.wheel_category, made=excluded.made,
                 source=excluded.source, shopping_list=excluded.shopping_list""",
            d,
        )
    # Rebuild all ingredients from the shopping lists.
    conn.execute("DELETE FROM dish_ingredients")  # empty the table (no WHERE = every row)
    for dish_id, text in conn.execute("SELECT id, shopping_list FROM dishes").fetchall():  # each dish's id + list
        conn.executemany(  # run the same INSERT once per item in the list below
            "INSERT INTO dish_ingredients (dish_id, ingredient, keyword, alt_group) VALUES (?, ?, ?, ?)",
            [(dish_id, raw, kw, grp) for raw, kw, grp in parse_ingredients(text)],  # one tuple per ingredient
        )


# ------------------------------------------------------------ main
def main():
    try:
        wb = load_workbook(XLSX, data_only=True, read_only=True)  # open the xlsx; formulas as values; read-only is faster
    except PermissionError:                                        # file locked by Excel
        sys.exit("Couldn't open the spreadsheet - close it in Excel and try again.")
    except FileNotFoundError:                                      # file missing
        sys.exit(f"Spreadsheet not found: {XLSX}")

    conn = db.connect()  # open pantry.db (creates tables if needed)
    with conn:  # one transaction: commits at the end, or rolls back on any error
        new_items = import_items(conn, wb["Pantry"])   # load pantry items
        import_dishes(conn, collect_dishes(wb))        # load dishes + ingredients
    wb.close()  # done with the spreadsheet

    print(f"New pantry items added: {new_items}")                    # report new items
    for table in ("items", "dishes", "dish_ingredients"):            # each table...
        n = conn.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0]  # ...count its rows
        print(f"{table:17} {n} rows")                                # padded to line up

    print("\nSample ingredient parsing:")  # show a few dishes so you can eyeball the keywords
    for name in ("Egg white frittata", "Avocado toast", "Black Bean Quesadillas"):
        # JOIN links each ingredient row to its dish row via dish_id = dishes.id.
        rows = conn.execute(
            """SELECT i.alt_group, i.keyword FROM dish_ingredients i
               JOIN dishes d ON d.id = i.dish_id WHERE d.name = ? ORDER BY i.id""", (name,))
        groups = {}                                       # alt_group -> list of keywords
        for grp, kw in rows:                              # each ingredient row
            groups.setdefault(grp, []).append(kw)         # add keyword to its group
        print(f"  {name}: " + " | ".join(" OR ".join(g) for g in groups.values()))  # 'a | b OR c | d'
    conn.close()  # close the database


if __name__ == "__main__":  # only when run directly (not when imported)
    main()
