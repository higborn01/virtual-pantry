"""Database layer: opens pantry.db, creates the tables, and holds shared helpers."""
import os       # file-path tools (joining folder + file names)
import sqlite3  # Python's built-in SQLite database library

HERE = os.path.dirname(os.path.abspath(__file__))  # the folder this file lives in (pantry_app/)
DB_PATH = os.path.join(HERE, "pantry.db")          # full path of the database file, next to this file

# SQL that builds the three tables. One big string, run all at once in connect().
# "IF NOT EXISTS" makes this safe to run every time the app starts.
# COLLATE NOCASE means 'Rice' and 'rice' count as the same name for UNIQUE checks.
SCHEMA = """
-- One row per thing in the pantry / fridge / freezer.
CREATE TABLE IF NOT EXISTS items (
  id INTEGER PRIMARY KEY,                    -- auto-numbered row id (1, 2, 3, ...)
  name TEXT NOT NULL UNIQUE COLLATE NOCASE,  -- item name; required; no two items share a name
  category TEXT NOT NULL DEFAULT 'Other',    -- shelf it goes on, e.g. 'Spices'
  qty REAL,                                  -- how many on hand; REAL = decimal number; NULL = unknown
  unit TEXT DEFAULT '',                      -- 'jar', 'bag', ... ('' if none)
  min_qty REAL DEFAULT 0,                    -- below this counts as 'Low'
  location TEXT NOT NULL DEFAULT 'Pantry',   -- Pantry, Fridge, Freezer or Counter
  updated DATE,                              -- last change, 'YYYY-MM-DD'; NULL = unknown
  notes TEXT DEFAULT ''                      -- free text
);

-- One row per dish from Meal History and/or Dinner Wheel.
CREATE TABLE IF NOT EXISTS dishes (
  id INTEGER PRIMARY KEY,                    -- auto-numbered row id
  name TEXT NOT NULL UNIQUE COLLATE NOCASE,  -- dish name; no duplicates (case ignored)
  category TEXT,                             -- Breakfast, Dinner, Side, ... (Meal History)
  cuisine TEXT,                              -- American, Italian, ...
  status TEXT,                               -- e.g. 'Made before (past chats)'
  notes TEXT DEFAULT '',                     -- free text
  wheel_category TEXT,                       -- Weeknight, Pasta Night, ... (Dinner Wheel)
  made INTEGER DEFAULT 0,                    -- 1 = we've made it, 0 = not yet
  source TEXT,                               -- 'Meal History', 'Dinner Wheel' or 'Both'
  shopping_list TEXT DEFAULT ''              -- original comma-separated text, for display
);

-- One row per ingredient of a dish (split out of the shopping list).
CREATE TABLE IF NOT EXISTS dish_ingredients (
  id INTEGER PRIMARY KEY,                    -- auto-numbered row id
  dish_id INTEGER NOT NULL REFERENCES dishes(id) ON DELETE CASCADE,  -- which dish; deleting the dish deletes these rows too
  ingredient TEXT NOT NULL,                  -- raw text, e.g. 'feta or mozzarella'
  keyword TEXT NOT NULL,                     -- normalized, e.g. 'feta' (used for matching)
  alt_group INTEGER NOT NULL                 -- rows with the same dish_id + alt_group are alternatives
);

-- Extra things to buy that aren't pantry items (e.g. missing recipe ingredients).
CREATE TABLE IF NOT EXISTS shopping_extras (
  id INTEGER PRIMARY KEY,                    -- auto-numbered row id (also keeps them in the order added)
  name TEXT NOT NULL UNIQUE COLLATE NOCASE,  -- what to buy, e.g. '1 can black beans'
  added DATE                                 -- when it was added, 'YYYY-MM-DD'
);
"""


def connect(path=DB_PATH):
    """Open the database (creating it if needed) and make sure the tables exist."""
    conn = sqlite3.connect(path)              # open the file (SQLite creates it if missing)
    conn.row_factory = sqlite3.Row            # rows behave like dicts: row["name"]
    conn.execute("PRAGMA foreign_keys = ON")  # SQLite only enforces REFERENCES when this is on
    conn.executescript(SCHEMA)                # run all the CREATE TABLE statements above
    return conn                               # hand the open connection back to the caller


def status_of(qty, min_qty):
    """'OK', 'Low', 'Out', or '?' - same rule as the spreadsheet's Status column."""
    if qty is None:            # quantity never entered
        return "?"             #   -> unknown
    if qty <= 0:               # none left
        return "Out"
    if qty < (min_qty or 0):   # below the minimum (a blank minimum counts as 0)
        return "Low"
    return "OK"                # otherwise we have enough
