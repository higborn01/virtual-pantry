"""Virtual Pantry - illustrated pantry / fridge / freezer, backed by pantry.db.

Run:  python pantry_app.py   (or double-click run_pantry.bat)
First time: python import_excel.py  (loads Meal_Plan_and_History.xlsx into pantry.db)
Needs: Python 3 with tkinter + openpyxl (both come with Anaconda).

We'll build this in steps - see CLAUDE.md. Each step fills in one TODO below.
"""
import datetime as dt                 # dates (leftover ages, 'updated' stamps) - used in later steps
import math                           # math.ceil (round up) for counting tile rows
import os                             # file paths (for the window icon)
import random                         # random positions for the background snowflakes
import tkinter as tk                  # Python's built-in GUI toolkit (windows, canvas, buttons)
import tkinter.font as tkfont         # Font objects - can measure how wide a piece of text will be
from tkinter import ttk, messagebox   # ttk = nicer-looking widgets; messagebox = pop-up dialogs

from tkinter import simpledialog  # small pop-up that asks for one line of text

import cook_match            # 'What can I cook?' matching (Step 7)
import db                    # our database module (db.py)
import export_excel          # 'Export to Excel' (Step 5)
from emoji_icons import emoji_for, emoji_image  # picks each item's emoji + draws it in colour
from db import status_of     # the OK / Low / Out / ? rule, used by later steps

# ----------------------------------------------------------------- settings
LOCATIONS = ["Pantry", "Fridge", "Freezer", "Counter"]  # where an item can live
LEFTOVER_DAYS = 4  # leftovers older than this get flagged

# ----------------------------------------------------- STEP 3: look & feel
OUTLINE = "#1f2d3d"   # the thick dark cartoon outline used on almost everything

# Shelf order inside each unit (categories not in this list go at the end).
CATEGORIES = ["Produce", "Dairy & Eggs", "Frozen", "Canned & Dry", "Grains & Pasta", "Baking",
              "Spices", "Sauces & Condiments", "Bread & Tortillas", "Snacks", "Leftovers", "Other"]

CAT_COLORS = {                      # bright, flat tile colour for each category
    "Produce": "#7ed957",           # leafy green
    "Dairy & Eggs": "#fff1a8",      # butter yellow
    "Frozen": "#9be7ff",            # ice blue
    "Canned & Dry": "#ffa36c",      # orange
    "Grains & Pasta": "#ffd166",    # golden
    "Baking": "#ffc2e2",            # bubblegum pink
    "Spices": "#ff7f5c",            # paprika
    "Sauces & Condiments": "#ff6b81",  # ketchup red-pink
    "Bread & Tortillas": "#e9c46a", # toast
    "Snacks": "#c99bff",            # purple
    "Leftovers": "#8ec5ff",         # tupperware blue
    "Other": "#cfd8dc",             # grey
}

STATUS_COLORS = {"OK": "#46d160", "Low": "#ffbe0b", "Out": "#ff4d5e", "?": "#b0bac5"}  # badge colours

# One colour theme per storage unit: outer body, trim (signs/feet), inside wall,
# shelf boards, shelf-label text, and the soft shadow tiles cast on the inside wall.
THEMES = {
    "Pantry":  dict(body="#d98c4a", trim="#9a5a2a", inside="#f6d6a3", shelf="#b8692f", label="#6b3a14", shadow="#d9b27c"),
    "Freezer": dict(body="#dff3ff", trim="#5aa9e6", inside="#c7e8ff", shelf="#8fd0ff", label="#1d5c8f", shadow="#a6cdea"),
    "Fridge":  dict(body="#ffffff", trim="#4fb3e8", inside="#e6f5ff", shelf="#a8dcff", label="#1d5c8f", shadow="#bcd9ee"),
    "Counter": dict(body="#f5e6cf", trim="#c48a5a", inside="#fff8ec", shelf="#e3c9a4", label="#7a4b25", shadow="#e5d3b8"),
}
COLD = ("Fridge", "Freezer")  # these get a door handle on the right

# Columns drawn left to right; each column stacks one or more compartments top to bottom.
COLUMNS = [("Pantry", ["Counter", "Pantry"]), ("Fridge", ["Freezer", "Fridge"])]  # counter sits on top of the pantry

# Sizes in pixels
MARGIN = 24        # space between the canvas edge and the units
COL_GAP = 28       # space between units
TOP = 44           # y where the units start (room for the sign above)
DOOR_GAP = 18      # gap between the freezer and fridge doors
TILE_MIN_W = 104   # tiles are at least this wide (they stretch to fill the row)
EMOJI_PX = 28      # emoji picture size on a tile (tile height is worked out from the font size in build_layout)
GAP = 10           # space between tiles
FEET = 14          # height of the little feet under each unit
R_UNIT = 46        # corner roundness of the units (bigger = more bubbly)
R_WALL = 34        # corner roundness of the inside back walls
R_TILE = 22        # corner roundness of the item tiles

# Draw order, bottom to top. Every canvas shape gets one of these tags.
LAYERS = ("bg", "body", "inside", "board", "tile")


def mix(c1, c2, t):
    """Blend two '#rrggbb' colours: t=0 gives c1, t=1 gives c2."""
    a = [int(c1[i:i + 2], 16) for i in (1, 3, 5)]   # c1 as [r, g, b] numbers 0-255
    b = [int(c2[i:i + 2], 16) for i in (1, 3, 5)]   # c2 as [r, g, b]
    return "#" + "".join(f"{round(x + (y - x) * t):02x}" for x, y in zip(a, b))  # blend each channel, back to hex


def round_rect(canvas, x1, y1, x2, y2, r, **kw):
    """Draw a rounded rectangle. tkinter has none, so we trace one as a polygon with quarter-circle corners."""
    r = max(0, min(r, (x2 - x1) / 2, (y2 - y1) / 2))  # corner radius can't be bigger than half the box
    pts = []                                          # list of x, y, x, y, ... around the outline
    corners = [(x2 - r, y1 + r, -90), (x2 - r, y2 - r, 0),  # each corner's circle centre + starting angle:
               (x1 + r, y2 - r, 90), (x1 + r, y1 + r, 180)]  # top-right, bottom-right, bottom-left, top-left
    for cx, cy, start in corners:                     # go clockwise round the four corners
        for step in range(7):                         # 7 points per corner = a smooth quarter circle
            a = math.radians(start + step * 15)       # angle of this point (0, 15, ... 90 degrees past the start)
            pts += [cx + r * math.cos(a), cy + r * math.sin(a)]  # point on the circle (cos = across, sin = down)
    return canvas.create_polygon(pts, **kw)           # straight lines between the points join the corners up


def qty_label(item):
    """'1 jar', '0.5 bag', '? jar' (unknown qty)."""
    q = item["qty"]                                  # number or None
    num = "?" if q is None else f"{q:g}"             # :g drops the '.0' (1.0 -> '1', 0.2 -> '0.2')
    return f"{num} {item['unit'] or ''}".strip()     # add the unit; strip removes a trailing space if no unit


# ------------------------------------------------------- STEP 1: data layer
class Store:
    """Holds the pantry items and dishes in memory, loaded from pantry.db."""

    def __init__(self):
        self.conn = db.connect()  # open the database (keeps it open while the app runs)
        self.items = []           # one dict per pantry item
        self.dishes = []          # one dict per dish
        self.load()               # fill the two lists

    def load(self):
        # Read every row; dict(r) turns each sqlite3.Row into a plain dict.
        self.items = [dict(r) for r in self.conn.execute("SELECT * FROM items ORDER BY name")]
        self.dishes = [dict(r) for r in self.conn.execute("SELECT * FROM dishes ORDER BY name")]

        self.groups = {}                                 # dish id -> {alt_group: [(raw text, keyword), ...]} for 'What can I cook?'
        for r in self.conn.execute("SELECT dish_id, alt_group, ingredient, keyword FROM dish_ingredients ORDER BY id"):
            self.groups.setdefault(r["dish_id"], {}).setdefault(r["alt_group"], []).append((r["ingredient"], r["keyword"]))

    # ---- STEP 5: every change is written to pantry.db straight away ----
    FIELDS = ("name", "category", "qty", "unit", "min_qty", "location", "updated", "notes")  # item columns we save

    def update(self, item):
        """Save an edited item."""
        sets = ", ".join(f"{f} = ?" for f in self.FIELDS)  # 'name = ?, category = ?, ...'
        self.conn.execute(f"UPDATE items SET {sets} WHERE id = ?",  # UPDATE = change an existing row; WHERE picks which one
                          [item[f] for f in self.FIELDS] + [item["id"]])  # the new values, then the id for WHERE
        self.conn.commit()                               # commit = make the change permanent on disk

    def insert(self, item):
        """Save a new item; the database gives it its id."""
        cols = ", ".join(self.FIELDS)                    # 'name, category, ...'
        marks = ", ".join("?" for _ in self.FIELDS)      # '?, ?, ...' - one placeholder per column
        cur = self.conn.execute(f"INSERT INTO items ({cols}) VALUES ({marks})", [item[f] for f in self.FIELDS])
        self.conn.commit()                               # make it permanent
        item["id"] = cur.lastrowid                       # lastrowid = the id SQLite just gave the new row
        self.items.append(item)                          # add it to the in-memory list too

    def delete(self, item):
        """Remove an item for good (the caller has already asked the user)."""
        self.conn.execute("DELETE FROM items WHERE id = ?", (item["id"],))  # DELETE = remove the matching row
        self.conn.commit()
        self.items.remove(item)                          # and from the list

    # ---- shopping-list extras (things to buy that aren't pantry items) ----
    def extras(self):
        return [r["name"] for r in self.conn.execute("SELECT name FROM shopping_extras ORDER BY id")]  # in the order added

    def add_extra(self, name):
        self.conn.execute("INSERT OR IGNORE INTO shopping_extras (name, added) VALUES (?, ?)",  # ignore if already listed
                          (name, dt.date.today().isoformat()))
        self.conn.commit()

    def remove_extra(self, name):
        self.conn.execute("DELETE FROM shopping_extras WHERE name = ?", (name,))
        self.conn.commit()


# ------------------------------------------------------------- STEP 2+: UI
class App(tk.Tk):  # the main window; tk.Tk is tkinter's top-level window class
    def __init__(self):
        super().__init__()                # set up the underlying tk window
        self.title("Virtual Pantry")      # window title bar text
        self.geometry("1400x850")         # starting size: width x height in pixels
        try:                              # the icon file might be missing - that's fine
            self.iconbitmap(os.path.join(db.HERE, "pantry.ico"))  # fridge icon in the title bar + taskbar
        except tk.TclError:               # tkinter raises TclError if it can't load the icon
            pass                          #   keep the default icon
        self.store = Store()              # load data from pantry.db
        self.selected = None              # the item currently clicked (none yet)
        self._redraw_job = None           # id of a pending delayed redraw (see schedule_redraw)
        self.build_layout()               # create the widgets
        self.draw_shelves()               # draw the shelves and tiles

    def build_layout(self):
        """Top bar across the top; shelf canvas on the left; tabs on the right."""
        # ---- fonts used on the canvas (Font objects can measure text width) ----
        self.fonts = {
            "sign": tkfont.Font(family="Segoe UI Black", size=13),          # unit signs: PANTRY, FRIDGE...
            "shelf": tkfont.Font(family="Segoe UI Black", size=10),         # shelf labels: 'Spices · 26'
            "tile": tkfont.Font(family="Segoe UI", size=8, weight="bold"),  # item name on a tile
            "small": tkfont.Font(family="Segoe UI", size=8),                # qty on a tile
            "badge": tkfont.Font(family="Segoe UI Black", size=7),          # OK / Low / Out badge
        }
        self.line_h = self.fonts["tile"].metrics("linespace")  # height of one line of tile text (bigger on high-DPI screens)
        self.tile_h = 46 + 2 * self.line_h + 30          # tile height: emoji area + 2 name lines + qty/badge row

        # ---- top bar ----
        bar = ttk.Frame(self, padding=(10, 8))           # a box to hold the top-bar widgets; padding = space inside its edges
        bar.pack(side="top", fill="x")                   # put it at the top, stretched across the full width

        ttk.Label(bar, text="Virtual Pantry", font=("Segoe UI Black", 18)).pack(side="left")  # big chunky title on the left

        ttk.Label(bar, text="Search:").pack(side="left", padx=(30, 4))  # small label; padx=(left, right) spacing
        self.search_var = tk.StringVar()                 # a StringVar holds the text box's contents; we can watch it for changes
        self.search = ttk.Entry(bar, textvariable=self.search_var, width=30)  # the search box, linked to search_var
        self.search.pack(side="left")                    # place it next to the "Search:" label
        self.search_var.trace_add("write", lambda *args: self.schedule_redraw())  # every time the text changes, redraw the shelves

        ttk.Button(bar, text="+ Add item", command=self.add_item).pack(side="left", padx=10)  # button; command = function run on click
        ttk.Button(bar, text="Export to Excel", command=self.export_excel).pack(side="left")  # write items to the spreadsheet (Step 5)

        self.counts_label = ttk.Label(bar, font=("Segoe UI", 11))  # label for "144 items · 2 low · 15 out"; text set below
        self.counts_label.pack(side="right")             # stick it to the right edge
        self.update_counts()                             # fill in the counts now

        # ---- main area: split into left (shelves) and right (tabs) ----
        split = ttk.PanedWindow(self, orient="horizontal")  # PanedWindow = side-by-side panes with a draggable divider
        split.pack(side="top", fill="both", expand=True)    # fill the rest of the window; expand = take any spare space

        # left pane: a canvas (drawing surface) with a vertical scrollbar
        left = ttk.Frame(split)                          # box to hold the canvas + scrollbar
        self.canvas = tk.Canvas(left, background="#e6f6ff", highlightthickness=0)  # the drawing area; pale icy sky; no focus border
        scroll = ttk.Scrollbar(left, orient="vertical", command=self.canvas.yview)  # dragging the scrollbar scrolls the canvas up/down
        self.canvas.configure(yscrollcommand=scroll.set)  # and scrolling the canvas moves the scrollbar thumb (the two talk both ways)
        scroll.pack(side="right", fill="y")              # scrollbar on the right edge, full height
        self.canvas.pack(side="left", fill="both", expand=True)  # canvas fills the rest
        self.canvas.bind("<MouseWheel>", self.on_mousewheel)      # mouse wheel over the canvas -> scroll it
        self.canvas.bind("<Configure>", lambda event: self.schedule_redraw())  # <Configure> fires when the canvas is resized -> redraw
        # tag_bind = react to the mouse on canvas shapes with a given tag (stays active across redraws)
        self.canvas.tag_bind("tile", "<Button-1>", self.on_tile_click)          # single left-click on a tile -> select it
        self.canvas.tag_bind("tile", "<Double-Button-1>", self.on_tile_double)  # double-click -> use one up (-1)
        self.canvas.tag_bind("tile", "<Enter>", lambda e: self.canvas.configure(  # mouse moves onto a tile...
            cursor="hand2" if self.item_under_mouse() else ""))                   # ...show the pointing-hand cursor
        self.canvas.tag_bind("tile", "<Leave>", lambda e: self.canvas.configure(cursor=""))  # mouse leaves -> normal cursor
        split.add(left, weight=3)                        # add the left pane; weight = how much extra width it gets (3 parts)

        # right pane: a Notebook = a set of tabs
        self.tabs = ttk.Notebook(split, width=380)       # the tab container, about 380 px wide to start
        split.add(self.tabs, weight=1)                   # add the right pane (1 part of extra width)
        self.tab = {}                                    # dict: tab name -> its frame, so later steps can fill each tab
        for name in ["Item", "Shopping list", "What can I cook?", "Leftovers"]:  # each tab
            frame = ttk.Frame(self.tabs, padding=12)     # one page of the notebook
            self.tabs.add(frame, text=name)              # add the page with its tab title
            self.tab[name] = frame                       # remember it by name
        style = ttk.Style(self)                          # shared look for the lists in the tabs
        style.configure("Treeview", rowheight=self.line_h + 12, font=("Segoe UI", 10))  # taller rows that fit the text
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))              # bold column titles
        self.build_item_panel()                          # fill the Item tab (Step 4)
        self.build_shopping_tab()                        # Step 6
        self.build_cook_tab()                            # Step 7
        self.build_leftovers_tab()                       # Step 8
        self.tabs.bind("<<NotebookTabChanged>>", lambda e: self.refresh_current_tab())  # virtual event: fires when you switch tabs

    def update_counts(self):
        """Refresh the 'N items · N low · N out' label in the top bar."""
        statuses = [status_of(i["qty"], i["min_qty"]) for i in self.store.items]  # OK/Low/Out/? for every item
        self.counts_label.configure(                     # change the label's text
            text=f"{len(statuses)} items  ·  {statuses.count('Low')} low  ·  {statuses.count('Out')} out")

    def on_mousewheel(self, event):
        """Scroll the shelf canvas with the mouse wheel."""
        # On Windows one wheel 'click' gives event.delta = 120 (up) or -120 (down); yview_scroll wants -1 / +1
        self.canvas.yview_scroll(int(-event.delta / 120), "units")

    def add_item(self):
        """'+ Add item' button: make a blank item, select it, and put the cursor in its Name box."""
        names = {i["name"].lower() for i in self.store.items}  # names already in use (lower-case)
        name, n = "New item", 2                          # first choice of name, and a counter for 'New item 2', '3'...
        while name.lower() in names:                     # taken?
            name, n = f"New item {n}", n + 1             #   try the next number
        item = dict(name=name, category="Other", qty=1.0, unit="", min_qty=0.0,  # a fresh item...
                    location="Pantry", updated=dt.date.today().isoformat(), notes="")      # ...dated today
        self.store.insert(item)                          # save it (gets its id) and add it to the list
        self.select(item)                                # show it in the Item tab
        self.update_counts()                             # 1 more item
        self.name_entry.focus_set()                      # put the typing cursor in the Name box...
        self.name_entry.select_range(0, "end")           # ...with the text highlighted, so typing replaces it

    # ------------------------------------------------- STEP 4: item panel
    def build_item_panel(self):
        """Create the widgets in the Item tab (once). refresh_item_panel() fills them in."""
        tab = self.tab["Item"]                           # the Item tab's frame
        for widget in tab.winfo_children():              # remove the 'Coming in Step 4' placeholder
            widget.destroy()
        style = ttk.Style(self)                          # ttk styles = named looks shared by widgets
        style.configure("Big.TButton", font=("Segoe UI Black", 18), padding=(14, 0))  # chunky - / + buttons

        # Shown when nothing is selected:
        self.item_hint = ttk.Label(tab, foreground="grey", justify="left",
                                   text="Click an item on the shelves to see it here.\n\n"
                                        "Double-click a tile to use one up (-1).")
        # Shown when an item is selected:
        self.item_form = ttk.Frame(tab)                  # holds everything below

        head = ttk.Frame(self.item_form)                 # header row: emoji + name
        head.pack(fill="x")                              # full width
        self.item_pic = ttk.Label(head)                  # big emoji picture (image set in refresh)
        self.item_pic.pack(side="left", padx=(0, 10))    # on the left
        self.item_title = ttk.Label(head, font=("Segoe UI Black", 15), wraplength=280, justify="left")  # item name; wraps at 280 px
        self.item_title.pack(side="left", fill="x", expand=True)  # takes the rest of the row

        info = ttk.Frame(self.item_form)                 # status badge + where it lives
        info.pack(fill="x", pady=(8, 0))                 # a bit of space above
        self.item_status = tk.Label(info, font=("Segoe UI Black", 10), fg=OUTLINE, padx=10)  # coloured status badge (tk.Label lets us set bg colour)
        self.item_status.pack(side="left")
        self.item_where = ttk.Label(info, foreground="grey")  # 'Spices · Pantry'
        self.item_where.pack(side="left", padx=8)

        row = ttk.Frame(self.item_form)                  # the big - qty + row
        row.pack(pady=16)                                # centred, with space around it
        ttk.Button(row, text="−", width=2, style="Big.TButton", command=lambda: self.bump(-1)).pack(side="left")  # minus 1
        self.item_qty = ttk.Label(row, font=("Segoe UI Black", 20), anchor="center")  # current qty, big
        self.item_qty.pack(side="left", padx=16)
        ttk.Button(row, text="+", width=2, style="Big.TButton", command=lambda: self.bump(+1)).pack(side="left")  # plus 1

        form = ttk.Frame(self.item_form)                 # the editable fields, as a grid (label | box)
        form.pack(fill="x")
        form.columnconfigure(1, weight=1)                # column 1 (the boxes) stretches with the panel
        self.fields = {}                                 # field key -> StringVar holding its text
        specs = [("name", "Name"), ("qty", "Qty on hand"), ("unit", "Unit"), ("min_qty", "Min to keep"),
                 ("category", "Category"), ("location", "Location"), ("notes", "Notes")]  # (item key, label)
        for r, (key, label) in enumerate(specs):         # one grid row per field
            ttk.Label(form, text=label).grid(row=r, column=0, sticky="w", pady=3, padx=(0, 10))  # label, left-aligned (w = west)
            var = tk.StringVar()                         # holds the box's text
            if key == "category":                        # Combobox = text box with a drop-down list
                box = ttk.Combobox(form, textvariable=var, values=CATEGORIES)  # can also type a new category
            elif key == "location":
                box = ttk.Combobox(form, textvariable=var, values=LOCATIONS, state="readonly")  # readonly = pick from list only
            else:
                box = ttk.Entry(form, textvariable=var)  # plain one-line text box
            box.grid(row=r, column=1, sticky="ew", pady=3)  # sticky ew = stretch east-west to fill the column
            box.bind("<Return>", lambda e: self.apply_edits())  # pressing Enter in any box applies the changes
            self.fields[key] = var                       # remember the variable
            if key == "name":                            # keep the Name box handy for '+ Add item'
                self.name_entry = box

        btns = ttk.Frame(self.item_form)                 # action buttons
        btns.pack(fill="x", pady=(14, 0))
        ttk.Button(btns, text="Apply changes", command=self.apply_edits).pack(side="left")      # use what's typed in the boxes
        ttk.Button(btns, text="Used it up", command=self.used_it_up).pack(side="left", padx=6)  # qty -> 0
        ttk.Button(btns, text="Remove", command=self.remove_item).pack(side="right")            # delete (asks first)

        self.item_updated = ttk.Label(self.item_form, foreground="grey")  # 'Last updated: ...'
        self.item_updated.pack(anchor="w", pady=(12, 0))
        self.refresh_item_panel()                        # show the hint (nothing selected yet)

    def refresh_item_panel(self):
        """Show the selected item's details in the Item tab (or the hint if none)."""
        item = self.selected                             # the selected item, or None
        if item is None:                                 # nothing selected
            self.item_form.pack_forget()                 # pack_forget = hide a widget (it still exists)
            self.item_hint.pack(anchor="w")              # show the hint
            return
        self.item_hint.pack_forget()                     # hide the hint
        self.item_form.pack(fill="both", expand=True)    # show the details

        self.item_pic.configure(image=emoji_image(emoji_for(item), 48) or "")  # big emoji ('' = no image if Pillow missing)
        self.item_title.configure(text=item["name"])     # name
        status = status_of(item["qty"], item["min_qty"])  # OK / Low / Out / ?
        self.item_status.configure(text=status, bg=STATUS_COLORS[status])  # badge text + colour
        self.item_where.configure(text=f"{item['category']}  ·  {item['location']}")  # where it lives
        self.item_qty.configure(text=qty_label(item))    # big qty
        for key, var in self.fields.items():             # fill each box
            value = item[key]                            # the item's value for this field
            if key in ("qty", "min_qty"):                # numbers: show 1 not 1.0, blank for unknown
                var.set("" if value is None else f"{value:g}")
            else:
                var.set(value or "")                     # text ('' instead of None)
        self.item_updated.configure(text=f"Last updated: {item['updated'] or 'unknown'}")

    def item_under_mouse(self):
        """The item whose tile is under the mouse pointer, or None."""
        for tag in self.canvas.gettags("current"):       # 'current' = the canvas shape under the mouse
            if tag.startswith("item"):                   # our per-tile tag, e.g. 'item42'
                item_id = int(tag[4:])                   # the number after 'item'
                return next((i for i in self.store.items if i["id"] == item_id), None)  # find that item (None if gone)
        return None                                      # mouse is on a label/sign, not a tile

    def on_tile_click(self, event):
        item = self.item_under_mouse()                   # which item was clicked?
        if item:                                         # (ignore clicks on shelf labels / signs)
            self.select(item)

    def on_tile_double(self, event):
        item = self.item_under_mouse()                   # which item was double-clicked?
        if item:
            self.select(item)                            # make sure it's the selected one
            self.bump(-1)                                # use one up

    def change_item(self, item, **changes):
        """The ONE place items get changed: update it, date it, save it, refresh the screen."""
        item.update(changes)                             # apply the new values (dict.update)
        if "updated" not in changes and item["category"] != "Leftovers":  # leftovers keep the day they were made
            item["updated"] = dt.date.today().isoformat()  # stamp today's date, 'YYYY-MM-DD'
        self.store.update(item)                          # save to pantry.db right away
        self.refresh_after_change()                      # update everything on screen

    def refresh_after_change(self):
        """Redraw everything that shows item data."""
        self.refresh_item_panel()                        # Item tab
        self.update_counts()                             # low/out counts in the top bar
        self.refresh_current_tab()                       # whichever side tab is showing
        self.schedule_redraw()                           # shelves (tiles move if location/category changed)

    def export_excel(self):
        """'Export to Excel' button."""
        try:
            n = export_excel.export(self.store.items)    # write the Pantry sheet (backs up the file first)
        except PermissionError:                          # Excel has the file open (locked)
            messagebox.showerror("Spreadsheet is open", "Close Meal_Plan_and_History.xlsx in Excel, then try again.")
            return
        except Exception as e:                           # anything else: show it rather than crash
            messagebox.showerror("Export failed", f"Couldn't export:\n{e}")
            return
        messagebox.showinfo("Exported", f"Wrote {n} items to the Pantry sheet.\n"
                                        f"Backup of the old file: Meal_Plan_and_History.backup.xlsx")

    def bump(self, delta):
        """+1 / -1 on the selected item (never below 0; unknown counts as 0)."""
        item = self.selected
        if item:
            self.change_item(item, qty=max(0, (item["qty"] or 0) + delta))  # 'or 0' turns None into 0

    def used_it_up(self):
        if self.selected:
            self.change_item(self.selected, qty=0)       # 0 = Out

    def apply_edits(self):
        """Check the typed values and apply them to the selected item."""
        item = self.selected
        if not item:
            return
        f = {key: var.get().strip() for key, var in self.fields.items()}  # every box's text, trimmed
        if not f["name"]:                                # name is required
            messagebox.showerror("Name needed", "Please give the item a name.")
            return
        if any(o is not item and o["name"].lower() == f["name"].lower() for o in self.store.items):  # another item has it?
            messagebox.showerror("Name taken", f"There's already an item called \"{f['name']}\".")
            return
        try:                                             # numbers might be typed wrong
            qty = float(f["qty"]) if f["qty"] else None  # blank qty = unknown
            min_qty = float(f["min_qty"]) if f["min_qty"] else 0.0  # blank min = 0
        except ValueError:                               # float('abc') raises ValueError
            messagebox.showerror("Not a number", "Qty and Min to keep must be numbers, like 2 or 0.5.")
            return
        self.change_item(item, name=f["name"], qty=qty, unit=f["unit"], min_qty=min_qty,
                         category=f["category"] or "Other", location=f["location"] or "Pantry", notes=f["notes"])

    def remove_item(self):
        """Delete the selected item - after asking."""
        item = self.selected
        if not item:
            return
        if not messagebox.askyesno("Remove item", f"Remove \"{item['name']}\" from the pantry?"):  # Yes/No pop-up
            return                                       # 'No' -> do nothing
        self.store.delete(item)                          # delete from pantry.db and the list
        self.selected = None                             # nothing selected now
        self.refresh_after_change()                      # hint, counts, tabs, shelves

    # ------------------------------------------------------ STEP 3: shelves
    def schedule_redraw(self):
        """Redraw shortly, not instantly - dragging the window edge fires dozens of resize events."""
        if self._redraw_job:                             # a redraw is already waiting?
            self.after_cancel(self._redraw_job)          #   cancel it...
        self._redraw_job = self.after(40, self.draw_shelves)  # ...and wait 40 ms from now instead (after = run later)

    def draw_shelves(self):
        """Draw the whole scene: sky, units, shelves, tiles."""
        self._redraw_job = None                          # the waiting redraw is now running
        c = self.canvas                                  # short name
        c.delete("all")                                  # wipe everything drawn last time
        W = c.winfo_width()                              # current canvas width in pixels
        if W < 100:                                      # window not shown yet (width is 1 at start-up)
            return                                       #   nothing sensible to draw
        query = self.search_var.get().strip().lower()    # search text, lower-case ('' = no search)

        by_loc = {}                                      # location -> list of items there
        for item in self.store.items:                    # every item
            loc = item["location"] if item["location"] in THEMES else "Pantry"  # unknown locations go in the pantry
            by_loc.setdefault(loc, []).append(item)      # add to that location's list

        # Compartments to show in each column: all of them, except the Counter when nothing is on it.
        columns = [(name, [loc for loc in locs if loc != "Counter" or by_loc.get("Counter")]) for name, locs in COLUMNS]
        # Wider columns for units holding more items (at least 15 'items' worth so small ones aren't squashed).
        weights = [max(15, sum(len(by_loc.get(loc, [])) for loc in locs)) for _, locs in columns]
        usable = W - 2 * MARGIN - COL_GAP * (len(columns) - 1)  # width left for the units themselves
        widths = self.column_widths(weights, usable, TILE_MIN_W + 64)  # every unit fits at least one full tile

        # Pass 1: draw the shelves + tiles in each compartment, and note where each compartment ends.
        comps = []                                       # [location, x0, y0, x1, y_bottom, is_last_in_column]
        x = MARGIN                                       # left edge of the current column
        for (name, locs), col_w in zip(columns, widths):  # each column with its width
            y = TOP                                      # start at the top
            for i, loc in enumerate(locs):               # each compartment in the column (freezer, then fridge)
                bottom = self.draw_compartment_contents(loc, by_loc.get(loc, []), x, y, x + col_w, query)
                comps.append([loc, x, y, x + col_w, bottom, i == len(locs) - 1])  # remember its box
                y = bottom + DOOR_GAP                    # next compartment starts below this one
            x += col_w + COL_GAP                         # move right to the next column

        # Pass 2: all units stand on the same snowy floor, so stretch each column's last compartment down to it.
        floor_y = max(comp[4] for comp in comps) + 20 + FEET  # floor = below the tallest unit, plus room for feet
        for comp in comps:                               # each compartment
            if comp[5]:                                  # last one in its column?
                comp[4] = floor_y - FEET                 #   stretch it down to just above the feet
            self.draw_compartment_body(*comp)            # draw its outer body, inside wall, sign (and feet)

        H = max(floor_y + 80, c.winfo_height())          # total drawing height (at least fill the visible area)
        self.draw_background(W, H, floor_y)              # sky, snowflakes, snow floor

        for layer in LAYERS:                             # put the layers in order: each raise moves that layer to the top,
            c.tag_raise(layer)                           #   so raising bottom->top leaves 'tile' on top of everything
        c.configure(scrollregion=(0, 0, W, H))           # how far the scrollbar can scroll

    @staticmethod                                        # doesn't use self - just maths
    def column_widths(weights, usable, min_w):
        """Share `usable` pixels by weight, but never give a column less than min_w."""
        fixed = {}                                       # column index -> width forced to the minimum
        while True:                                      # repeat until no column is too thin
            rest = [i for i in range(len(weights)) if i not in fixed]  # columns still shared by weight
            if not rest:                                 # every column is at the minimum (very narrow window)
                return [min_w] * len(weights)            #   just use the minimum for all
            free = usable - sum(fixed.values())          # pixels left after the fixed ones
            total = sum(weights[i] for i in rest)        # total weight of the rest
            widths = [fixed.get(i, free * weights[i] / total) for i in range(len(weights))]  # fixed or proportional share
            thin = [i for i in rest if widths[i] < min_w]  # any too thin?
            if not thin:                                 # no - done
                return widths
            for i in thin:                               # yes - pin them to the minimum and share again
                fixed[i] = min_w

    def draw_compartment_contents(self, loc, items, x0, y0, x1, query):
        """Draw the labelled shelves and tiles for one compartment. Returns the y where it ends."""
        c, theme = self.canvas, THEMES[loc]              # canvas + this unit's colours
        left = x0 + 28                                   # tiles start 28 px inside the body (clear of the round corners)
        right = x1 - (44 if loc in COLD else 28)         # fridge/freezer leave extra room on the right for the handle
        inner = right - left                             # width available for tiles
        per_row = max(1, int((inner + GAP) // (TILE_MIN_W + GAP)))  # how many minimum-width tiles fit in a row
        tile_w = (inner - GAP * (per_row - 1)) / per_row  # stretch the tiles so the row is exactly full
        y = y0 + 44                                      # first shelf starts below the sign and the rounded top corners

        # Group items by category, in CATEGORIES order (unknown categories last, alphabetically).
        order = {cat: i for i, cat in enumerate(CATEGORIES)}   # category -> position number
        groups = {}                                      # category -> its items
        for item in sorted(items, key=lambda it: it["name"].lower()):  # items A-Z
            groups.setdefault(item["category"], []).append(item)        # add to its category's list
        shelves = sorted(groups.items(), key=lambda kv: (order.get(kv[0], 99), kv[0]))  # shelves in order

        if not shelves:                                  # nothing stored here
            c.create_text((left + right) / 2, y + 30, text="(empty)", font=self.fonts["shelf"],
                          fill=theme["label"], tags="tile")  # say so
            y += 70                                      # leave some height
            self.draw_board(left, right, y, theme)       # one empty shelf board
            return y + 30                                # compartment ends here

        for cat, cat_items in shelves:                   # one shelf per category
            c.create_text(left, y, text=f"{cat}  ·  {len(cat_items)}", anchor="nw",  # label, top-left anchored
                          font=self.fonts["shelf"], fill=theme["label"], tags="tile")
            y += 22                                      # tiles start below the label
            for i, item in enumerate(cat_items):         # each item on this shelf
                row, col = divmod(i, per_row)            # which row / column it goes in (divmod = whole division + remainder)
                tx = left + col * (tile_w + GAP)         # tile's left x
                ty = y + row * (self.tile_h + GAP)            # tile's top y
                hit = not query or query in item["name"].lower() or query in item["category"].lower()  # matches the search?
                self.draw_tile(item, tx, ty, tile_w, theme, dimmed=not hit)  # draw it (faded if it doesn't match)
            rows = math.ceil(len(cat_items) / per_row)   # rows used (rounded up)
            y += rows * (self.tile_h + GAP) - GAP             # move down past the tiles (no gap after the last row)
            self.draw_board(left, right, y, theme)       # the shelf board the tiles sit on
            y += 12 + 20                                 # board thickness + space before the next shelf
        return y                                         # compartment ends here

    def draw_board(self, left, right, y, theme):
        """A chunky cartoon shelf board whose top edge is at y."""
        c = self.canvas                                  # short name
        round_rect(c, left - 10, y + 5, right + 10, y + 19, 7, fill=theme["shadow"], outline="", tags="board")  # shadow under it
        round_rect(c, left - 10, y, right + 10, y + 14, 7, fill=theme["shelf"], outline=OUTLINE, width=2.5, tags="board")  # the board - fully rounded ends (radius = half its height)
        c.create_line(left - 3, y + 3.5, right + 3, y + 3.5, fill=mix(theme["shelf"], "#ffffff", 0.5),  # light shine line on top
                      width=2, capstyle="round", tags="board")

    def draw_tile(self, item, x, y, w, theme, dimmed):
        """One item: rounded tile in its category colour, name, qty and a status badge."""
        c, h = self.canvas, self.tile_h                      # short names
        tags = ("tile", f"item{item['id']}")             # 'tile' layer + a per-item tag so Step 4 can find clicks
        status = status_of(item["qty"], item["min_qty"])  # OK / Low / Out / ?
        color = CAT_COLORS.get(item["category"], CAT_COLORS["Other"])  # tile fill colour
        line, text, badge = OUTLINE, OUTLINE, STATUS_COLORS[status]    # outline, text and badge colours
        if self.selected is item:                        # the selected tile gets a bright yellow glow ring behind it
            round_rect(c, x - 6, y - 6, x + w + 6, y + h + 6, R_TILE + 6, fill="#ffd60a", outline=OUTLINE, width=2.5, tags=tags)
        if dimmed:                                       # not matching the search: fade everything into the wall colour
            color, line, badge = (mix(col, theme["inside"], 0.75) for col in (color, OUTLINE, badge))
            text = line                                  # faded text too
        else:                                            # only solid tiles get a drop shadow
            round_rect(c, x + 3, y + 4, x + w + 3, y + h + 4, R_TILE, fill=theme["shadow"], outline="", tags=tags)
        round_rect(c, x, y, x + w, y + h, R_TILE, fill=color, outline=line, width=2.5, tags=tags)  # the tile body
        round_rect(c, x + 14, y + 5, x + w - 14, y + 15, 5, fill=mix(color, "#ffffff", 0.5), outline="", tags=tags)  # glossy shine strip (inset from the round corners)

        ey = y + 27                                      # centre height of the emoji
        c.create_oval(x + w / 2 - 18, ey - 18, x + w / 2 + 18, ey + 18,  # soft light bubble behind the emoji
                      fill=mix(color, "#ffffff", 0.55), outline="", tags=tags)
        pic = emoji_image(emoji_for(item), EMOJI_PX, faded=dimmed)  # the item's emoji as a colour picture (pale if dimmed; None if unavailable)
        if pic:                                          # got one?
            c.create_image(x + w / 2, ey, image=pic, tags=tags)  # place it centred in the bubble

        name = item["name"]                              # full name to start with
        label = c.create_text(x + w / 2, y + 46, text=name, width=w - 14, justify="center", anchor="n",  # top-centred under the emoji
                              font=self.fonts["tile"], fill=text, tags=tags)  # width = where the text wraps
        while len(name) > 4:                             # shorten until it fits in 2 lines
            x1, y1, x2, y2 = c.bbox(label)               # bbox = the box the text actually takes up
            if y2 - y1 <= 2 * self.line_h + 2:           # 2 lines or fewer? done
                break
            name = name[:-3].rstrip()                    # chop 3 characters off the end (and any trailing space)
            c.itemconfigure(label, text=name + "…")      # show the shorter name with an ellipsis
        c.create_text(x + 13, y + h - 14, text=qty_label(item), anchor="w",  # qty at bottom-left, nudged in from the round corner
                      font=self.fonts["small"], fill=text, tags=tags)

        bw = self.fonts["badge"].measure(status) + 14    # badge width = text width + padding
        bx1, by0, by1 = x + w - 11, y + h - 22, y + h - 7  # badge right, top, bottom (nudged in from the round corner)
        round_rect(c, bx1 - bw, by0, bx1, by1, 7, fill=badge, outline=line, width=1.5, tags=tags)  # pill shape
        c.create_text(bx1 - bw / 2, (by0 + by1) / 2, text=status, font=self.fonts["badge"], fill=text, tags=tags)  # badge text

    def draw_compartment_body(self, loc, x0, y0, x1, y1, is_last):
        """The unit around the shelves: shadow, feet, outer body, inside wall, handle and sign."""
        c, theme = self.canvas, THEMES[loc]              # canvas + colours
        round_rect(c, x0 + 7, y0 + 8, x1 + 7, y1 + 8, R_UNIT, fill="#9cc6e4", outline="", tags="body")  # soft shadow on the sky
        if is_last:                                      # bottom compartment of the column: give the unit feet
            for fx in (x0 + 26, x1 - 54):                # left foot, right foot
                round_rect(c, fx, y1 - 6, fx + 28, y1 + FEET, 10, fill=theme["trim"], outline=OUTLINE, width=2.5, tags="body")  # round little foot
        round_rect(c, x0, y0, x1, y1, R_UNIT, fill=theme["body"], outline=OUTLINE, width=3, tags="body")  # outer body
        c.create_line(x0 + R_UNIT, y0 + 5, x1 - R_UNIT, y0 + 5, fill=mix(theme["body"], "#ffffff", 0.6),  # shine along the top edge (between the round corners)
                      width=3, capstyle="round", tags="body")
        wall_right = x1 - (30 if loc in COLD else 10)    # inside wall stops short of the handle on cold units
        round_rect(c, x0 + 10, y0 + 10, wall_right, y1 - 10, R_WALL, fill=theme["inside"],  # inside back wall
                   outline=theme["trim"], width=2, tags="inside")
        if loc in COLD:                                  # door handle
            hy = y0 + 40                                 # handle top
            round_rect(c, x1 - 25, hy, x1 - 13, hy + min(110, y1 - y0 - 80), 6,  # tall rounded bar (shorter in small doors)
                       fill="#cfdbe7", outline=OUTLINE, width=2.5, tags="inside")

        # Sign straddling the top edge, with white text and a dark outline.
        cx, text = (x0 + x1) / 2, loc.upper()            # centre x, sign text
        sw = self.fonts["sign"].measure(text) + 34       # sign width from text width
        round_rect(c, cx - sw / 2, y0 - 16, cx + sw / 2, y0 + 18, 17, fill=theme["trim"], outline=OUTLINE, width=3, tags="tile")  # pill-shaped sign (radius = half its height)
        for dx, dy in ((-2, 0), (2, 0), (0, -2), (0, 2), (-1.5, -1.5), (1.5, 1.5), (-1.5, 1.5), (1.5, -1.5)):  # 8 directions
            c.create_text(cx + dx, y0 + 1 + dy, text=text, font=self.fonts["sign"], fill=OUTLINE, tags="tile")  # dark copies = outline
        c.create_text(cx, y0 + 1, text=text, font=self.fonts["sign"], fill="white", tags="tile")  # white text on top

    def draw_background(self, W, H, floor_y):
        """Icy-blue sky that fades lighter toward the ground, snowflakes, and a wavy snow floor."""
        c, bands = self.canvas, 18                       # canvas; number of colour bands in the sky
        for i in range(bands):                           # each horizontal band
            top, bottom = floor_y * i / bands, floor_y * (i + 1) / bands + 1  # its y range (+1 avoids hairline gaps)
            c.create_rectangle(0, top, W, bottom, fill=mix("#9fd8ff", "#e6f6ff", i / (bands - 1)),  # darker at top, lighter below
                               outline="", tags="bg")
        rng = random.Random(7)                           # fixed seed: same snowflakes every redraw (no flicker)
        for _ in range(int(W * floor_y / 7000)):         # more flakes for a bigger sky
            fx, fy, r = rng.uniform(0, W), rng.uniform(0, floor_y), rng.uniform(1.5, 3.5)  # position + size
            c.create_oval(fx - r, fy - r, fx + r, fy + r, fill="white", outline="", tags="bg")  # one flake
        pts, step = [], 90                               # wavy floor outline points; wave spacing
        for i, wx in enumerate(range(-step, int(W) + 2 * step, step)):  # from just off the left edge to past the right
            pts += [wx, floor_y + (-8 if i % 2 else 4)]  # alternate up/down = gentle bumps
        pts += [W + step, H + 10, -step, H + 10]         # then down and along the bottom to close the shape
        c.create_polygon(pts, smooth=True, fill="#ffffff", outline="#a9cfe8", width=3, tags="bg")  # the snow

    def select(self, item):
        """Make `item` the selected one: show it in the Item tab and highlight its tile."""
        self.selected = item                             # remember it
        self.tabs.select(self.tab["Item"])               # switch the right-hand side to the Item tab
        self.refresh_item_panel()                        # fill in its details
        self.schedule_redraw()                           # redraw so the glow moves to this tile

    # ------------------------------------------------- side tabs: shared helpers
    def refresh_current_tab(self):
        """Refill whichever side tab is showing (the Item tab refreshes itself)."""
        if not hasattr(self, "leftover_tree"):           # still building the window
            return
        name = self.tabs.tab(self.tabs.select(), "text")  # title of the selected tab
        refill = {"Shopping list": self.side_shopping_list,
                  "What can I cook?": self.side_what_can_i_cook,
                  "Leftovers": self.side_leftovers}.get(name)  # its refill method (None for Item)
        if refill:
            refill()

    def make_tree(self, parent, columns, headings, widths, show="tree headings"):
        """A Treeview (scrollable list/table) with a vertical scrollbar. columns = extra columns after the name."""
        box = ttk.Frame(parent)                          # holds the tree + scrollbar
        box.pack(fill="both", expand=True, pady=(8, 8))  # fills the tab
        tree = ttk.Treeview(box, columns=columns, show=show, selectmode="browse")  # browse = one row selected at a time
        bar = ttk.Scrollbar(box, orient="vertical", command=tree.yview)  # scrollbar <-> tree, both ways
        tree.configure(yscrollcommand=bar.set)
        bar.pack(side="right", fill="y")
        tree.pack(side="left", fill="both", expand=True)
        for col, title, width in zip(("#0",) + tuple(columns), headings, widths):  # '#0' = the built-in first (name) column
            tree.heading(col, text=title, anchor="w")    # column title
            tree.column(col, width=width, stretch=(col == "#0"), anchor="w")  # only the name column stretches
        return tree

    # ------------------------------------------------- STEP 6: shopping list
    def build_shopping_tab(self):
        tab = self.tab["Shopping list"]
        ttk.Label(tab, text="Everything Low or Out (not leftovers), plus extras.\nClick a row to tick it off.",
                  foreground="grey", justify="left").pack(anchor="w")
        self.shop_tree = self.make_tree(tab, (), ("Item",), (300,), show="tree")  # just one column, no headings
        self.shop_tree.bind("<ButtonRelease-1>", self.on_shop_click)  # click (on release) -> tick/untick
        self.shop_tree.tag_configure("cat", font=("Segoe UI Black", 10))  # category rows in a chunky font
        self.shop_tree.tag_configure("done", foreground="#9aa5b1")       # ticked rows greyed out
        self.ticked = set()                              # row ids that are ticked (kept while the app is open)

        add = ttk.Frame(tab)                             # 'add something' row
        add.pack(fill="x")
        self.extra_var = tk.StringVar()                  # text typed in the box
        entry = ttk.Entry(add, textvariable=self.extra_var)
        entry.pack(side="left", fill="x", expand=True)
        entry.bind("<Return>", lambda e: self.add_extra_typed())  # Enter = add
        ttk.Button(add, text="Add", command=self.add_extra_typed).pack(side="left", padx=(6, 0))

        btns = ttk.Frame(tab)
        btns.pack(fill="x", pady=(8, 0))
        ttk.Button(btns, text="Copy list", command=self.copy_shopping).pack(side="left")             # to the clipboard
        ttk.Button(btns, text="Remove ticked extras", command=self.clear_ticked_extras).pack(side="left", padx=6)
        self.shop_msg = ttk.Label(tab, foreground="grey")  # little confirmation message
        self.shop_msg.pack(anchor="w", pady=(6, 0))

    def shopping_rows(self):
        """[(category, [(row id, text), ...]), ...] for the shopping list."""
        order = {cat: i for i, cat in enumerate(CATEGORIES)}  # category -> position
        groups = {}                                      # category -> rows
        for item in sorted(self.store.items, key=lambda i: i["name"].lower()):  # A-Z
            status = status_of(item["qty"], item["min_qty"])
            if status in ("Low", "Out") and item["category"] != "Leftovers":  # needs buying (leftovers don't)
                text = f"{item['name']}   (have {qty_label(item)}, keep {item['min_qty']:g})"
                groups.setdefault(item["category"], []).append((f"item:{item['id']}", text))
        rows = sorted(groups.items(), key=lambda kv: (order.get(kv[0], 99), kv[0]))  # categories in shelf order
        extras = [(f"extra:{name}", name) for name in self.store.extras()]  # extras from recipes / typed
        if extras:
            rows.append(("Extras", extras))              # extras go last
        return rows

    def side_shopping_list(self):
        """Refill the shopping list."""
        tree = self.shop_tree
        tree.delete(*tree.get_children())                # clear it (* spreads the list into separate arguments)
        count = 0
        for cat, rows in self.shopping_rows():           # each category
            parent = tree.insert("", "end", text=f"{cat}  ·  {len(rows)}", open=True, tags=("cat",))  # '' = top level
            for iid, text in rows:                       # each thing to buy
                done = iid in self.ticked                # ticked off?
                tree.insert(parent, "end", iid=iid, text=("☑ " if done else "☐ ") + text, tags=("done",) if done else ())
                count += 1
        self.shop_msg.configure(text=f"{count} things on the list." if count else "Nothing to buy!")

    def on_shop_click(self, event):
        iid = self.shop_tree.identify_row(event.y)       # which row is at the click's height?
        if not iid or ":" not in iid:                    # empty space or a category row
            return
        self.ticked ^= {iid}                             # ^= toggles: adds it if missing, removes it if there
        self.side_shopping_list()                        # redraw with the new tick

    def add_extra_typed(self):
        text = self.extra_var.get().strip()              # what was typed
        if text:
            self.store.add_extra(text)                   # save it
            self.extra_var.set("")                       # clear the box
            self.side_shopping_list()

    def clear_ticked_extras(self):
        for iid in [i for i in self.ticked if i.startswith("extra:")]:  # ticked extras
            self.store.remove_extra(iid[len("extra:"):])  # delete from the database
            self.ticked.discard(iid)                     # and forget the tick
        self.side_shopping_list()

    def copy_shopping(self):
        """Copy everything NOT ticked to the clipboard, grouped by category."""
        lines = [f"Shopping list - {dt.date.today():%b %d}"]  # title with today's date, e.g. 'Nov 03'
        for cat, rows in self.shopping_rows():
            todo = [text.split("   (")[0] for iid, text in rows if iid not in self.ticked]  # names only, unticked only
            if todo:
                lines += ["", cat] + [f"  ☐ {t}" for t in todo]
        self.clipboard_clear()                           # empty the system clipboard
        self.clipboard_append("\n".join(lines))          # put the list on it
        self.shop_msg.configure(text="Copied! Paste it anywhere (Ctrl+V).")

    # ------------------------------------------------- STEP 7: what can I cook?
    def build_cook_tab(self):
        tab = self.tab["What can I cook?"]
        top = ttk.Frame(tab)
        top.pack(fill="x")
        ttk.Label(top, text="Show:").pack(side="left")
        self.cook_filter = tk.StringVar(value="All")     # which dish category to show
        cats = ["All", "Dinner", "Breakfast", "Side", "Snack", "Dessert", "Bread"]
        box = ttk.Combobox(top, textvariable=self.cook_filter, values=cats, state="readonly", width=12)
        box.pack(side="left", padx=6)
        box.bind("<<ComboboxSelected>>", lambda e: self.side_what_can_i_cook())  # refilter on change

        self.cook_tree = self.make_tree(tab, ("have", "missing"), ("Dish", "Have", "Missing"), (190, 100, 150))
        self.cook_tree.tag_configure("all", background="#d8f5dc")   # everything in stock -> green row
        self.cook_tree.tag_configure("most", background="#fff4cc")  # 70%+ -> yellow row
        self.cook_tree.bind("<<TreeviewSelect>>", lambda e: self.show_dish())  # row picked -> details below

        self.dish_title = ttk.Label(tab, font=("Segoe UI Black", 11), wraplength=360, justify="left")
        self.dish_title.pack(anchor="w")
        self.dish_info = ttk.Label(tab, wraplength=360, justify="left")  # shopping list + what's missing
        self.dish_info.pack(anchor="w", pady=(2, 6))
        self.dish_btn = ttk.Button(tab, text="Add missing to shopping list", command=self.add_missing)
        self.cook_results = {}                           # row id -> result dict

    def side_what_can_i_cook(self):
        """Score every dish against what's in stock and list them best first."""
        tree = self.cook_tree
        tree.delete(*tree.get_children())
        want = self.cook_filter.get()                    # 'All' or one category
        dishes = [d for d in self.store.dishes if want == "All" or d["category"] == want]
        self.cook_results = {}
        for r in cook_match.rank(dishes, self.store.groups, self.store.items):  # best first
            iid = str(r["dish"]["id"])                   # row id = dish id
            tag = "all" if r["pct"] == 100 else "most" if r["pct"] >= 70 else ""  # row colour
            tree.insert("", "end", iid=iid, text=r["dish"]["name"], tags=(tag,),
                        values=(f"{r['pct']}%  ({r['have']}/{r['total']})", ", ".join(r["missing"]) or "—"))
            self.cook_results[iid] = r
        self.show_dish()                                 # clear / refresh the details area

    def show_dish(self):
        sel = self.cook_tree.selection()                 # selected row ids (a tuple)
        r = self.cook_results.get(sel[0]) if sel else None
        if not r:                                        # nothing selected
            self.dish_title.configure(text="Pick a dish to see what's missing.")
            self.dish_info.configure(text="")
            self.dish_btn.pack_forget()
            return
        d = r["dish"]
        self.dish_title.configure(text=d["name"])
        info = f"Needs: {d['shopping_list']}"
        info += f"\n\nMissing: {', '.join(r['missing'])}" if r["missing"] else "\n\nYou have everything!"
        self.dish_info.configure(text=info)
        if r["missing"]:
            self.dish_btn.pack(anchor="w")               # show the button only when something is missing
        else:
            self.dish_btn.pack_forget()

    def add_missing(self):
        sel = self.cook_tree.selection()
        r = self.cook_results.get(sel[0]) if sel else None
        if r:
            for text in r["missing"]:                    # each missing ingredient
                self.store.add_extra(text)               # onto the shopping list (duplicates ignored)
            messagebox.showinfo("Added", f"Added {len(r['missing'])} things to the Shopping list.")

    # ------------------------------------------------- STEP 8: leftovers
    def build_leftovers_tab(self):
        tab = self.tab["Leftovers"]
        ttk.Label(tab, foreground="grey", justify="left",
                  text=f"Oldest first. Older than {LEFTOVER_DAYS} days = eat or toss!").pack(anchor="w")
        self.leftover_tree = self.make_tree(tab, ("age",), ("Leftover", "In since"), (180, 170))
        self.leftover_tree.tag_configure("old", background="#ffd6da", foreground="#a4161a")  # red rows
        self.leftover_tree.tag_configure("unknown", foreground="#7a8794")                   # grey rows
        self.leftover_tree.bind("<Double-1>", lambda e: self.leftover_selected() and self.select(self.leftover_selected()))  # double-click -> Item tab
        btns = ttk.Frame(tab)
        btns.pack(fill="x")
        ttk.Button(btns, text="Made today", command=self.leftover_made_today).pack(side="left")
        ttk.Button(btns, text="Finished", command=self.leftover_finished).pack(side="left", padx=6)
        ttk.Button(btns, text="+ Add leftover", command=self.add_leftover).pack(side="right")

    def side_leftovers(self):
        tree = self.leftover_tree
        tree.delete(*tree.get_children())
        today = dt.date.today()
        lefts = [i for i in self.store.items if i["category"] == "Leftovers"]
        lefts.sort(key=lambda i: i["updated"] or "")     # oldest date first; unknown dates ('') come first
        for item in lefts:
            if item["updated"]:                          # date known
                made = dt.date.fromisoformat(item["updated"])  # text -> date
                days = (today - made).days               # age in days (date - date = timedelta)
                age = f"{made:%b %d}  —  {days} day{'s' if days != 1 else ''}"  # e.g. 'Nov 01 — 2 days'
                tag = "old" if days > LEFTOVER_DAYS else ""
                if tag:
                    age += "  ⚠ eat or toss"
            else:
                age, tag = "date unknown (click Made today)", "unknown"
            tree.insert("", "end", iid=str(item["id"]), text=item["name"], values=(age,), tags=(tag,))

    def leftover_selected(self):
        sel = self.leftover_tree.selection()
        return next((i for i in self.store.items if str(i["id"]) == sel[0]), None) if sel else None

    def leftover_made_today(self):
        item = self.leftover_selected()
        if item:
            self.change_item(item, updated=dt.date.today().isoformat())  # age back to 0

    def leftover_finished(self):
        item = self.leftover_selected()
        if item and messagebox.askyesno("Finished", f"All gone? Remove \"{item['name']}\"?"):
            self.store.delete(item)
            if self.selected is item:
                self.selected = None
            self.refresh_after_change()

    def add_leftover(self):
        name = simpledialog.askstring("Add leftover", "What is it?", parent=self)  # pop-up text box (None if cancelled)
        if not name or not name.strip():
            return
        name = name.strip()
        names = {i["name"].lower() for i in self.store.items}
        base, n = name, 2
        while name.lower() in names:                     # make the name unique
            name, n = f"{base} ({n})", n + 1
        item = dict(name=name, category="Leftovers", qty=1.0, unit="container", min_qty=0.0,
                    location="Fridge", updated=dt.date.today().isoformat(), notes="")
        self.store.insert(item)                          # save it
        self.refresh_after_change()


if __name__ == "__main__":  # only when this file is run directly
    try:                                                 # Windows only; harmless to skip elsewhere
        import ctypes                                    # lets Python call Windows system functions
        ctypes.windll.shcore.SetProcessDpiAwareness(1)   # draw at the screen's real resolution (sharp, not blurry-stretched)
    except (AttributeError, OSError):                    # not Windows, or an old Windows without this function
        pass                                             #   just carry on
    App().mainloop()        # create the window and start tkinter's event loop (waits for clicks/keys)
