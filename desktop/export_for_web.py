"""Exports everything in pantry.db to pantry_export.json for the web app's one-time import.

Run:  python export_for_web.py
Then in the web app: menu (⋯) -> 'Import from desktop app' -> pick pantry_export.json.
The file holds your pantry data, so it stays here on your computer - don't put it in the website folder.
"""
import json   # write the data as JSON (text format the web app can read)
import os     # file paths

import db     # opens pantry.db

conn = db.connect()                                              # open the database
items = [dict(r) for r in conn.execute("SELECT * FROM items ORDER BY id")]  # every pantry item as a dict
extras = [r["name"] for r in conn.execute("SELECT name FROM shopping_extras ORDER BY id")]  # shopping extras

dishes = []                                                      # one entry per dish
for d in conn.execute("SELECT * FROM dishes ORDER BY name"):     # every dish, A-Z
    groups = {}                                                  # alt_group -> [[raw text, keyword], ...]
    for g in conn.execute("SELECT alt_group, ingredient, keyword FROM dish_ingredients WHERE dish_id = ? ORDER BY id",
                          (d["id"],)):                           # this dish's ingredients
        groups.setdefault(g["alt_group"], []).append([g["ingredient"], g["keyword"]])
    dishes.append(dict(name=d["name"], category=d["category"], made=d["made"],
                       shopping_list=d["shopping_list"], groups=list(groups.values())))  # ingredient groups in order

out = os.path.join(db.HERE, "pantry_export.json")                # save next to this script
with open(out, "w", encoding="utf-8") as f:                      # 'with' closes the file automatically when done
    json.dump(dict(version=1, items=items, extras=extras, dishes=dishes), f, ensure_ascii=False, indent=1)  # keep emoji/é readable
print(f"Saved {out}: {len(items)} items, {len(dishes)} dishes, {len(extras)} extras")  # report
