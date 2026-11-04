"""Creates a made-up sample spreadsheet (../Meal_Plan_and_History.xlsx) so anyone can try the app.

Run:  python make_sample_data.py   then   python import_excel.py   then   python pantry_app.py
"""
import datetime as dt                # dates for the leftovers
import os                            # file paths
from openpyxl import Workbook        # create .xlsx files

HERE = os.path.dirname(os.path.abspath(__file__))                        # this folder
OUT = os.path.join(os.path.dirname(HERE), "Meal_Plan_and_History.xlsx")  # the spreadsheet goes one folder up
day = lambda n: dt.date.today() - dt.timedelta(days=n)                   # the date n days ago

# (name, category, qty, unit, min, location, last updated, notes)
ITEMS = [
    ("Rice", "Grains & Pasta", 2, "bag", 1, "Pantry", None, ""),
    ("Spaghetti", "Grains & Pasta", 1, "box", 1, "Pantry", None, ""),
    ("Penne", "Grains & Pasta", 0, "box", 1, "Pantry", None, ""),
    ("Rolled oats", "Grains & Pasta", 1, "canister", 1, "Pantry", None, ""),
    ("Black beans", "Canned & Dry", 3, "cans", 2, "Pantry", None, ""),
    ("Chickpeas", "Canned & Dry", 1, "can", 2, "Pantry", None, "running low"),
    ("Diced tomatoes", "Canned & Dry", 2, "cans", 1, "Pantry", None, ""),
    ("Coconut milk", "Canned & Dry", 1, "can", 1, "Pantry", None, ""),
    ("All-purpose flour", "Baking", 1, "bag", 1, "Pantry", None, ""),
    ("Granulated sugar", "Baking", 1, "bag", 1, "Pantry", None, ""),
    ("Vanilla extract", "Baking", 1, "bottle", 0, "Pantry", None, ""),
    ("Ground cumin", "Spices", 1, "jar", 1, "Pantry", None, ""),
    ("Smoked paprika", "Spices", 0, "jar", 1, "Pantry", None, ""),
    ("Garlic powder", "Spices", 1, "jar", 1, "Pantry", None, ""),
    ("Salt", "Spices", 1, "box", 1, "Pantry", None, ""),
    ("Olive oil", "Sauces & Condiments", 1, "bottle", 1, "Pantry", None, ""),
    ("Soy sauce", "Sauces & Condiments", 1, "bottle", 1, "Fridge", None, ""),
    ("Marinara", "Sauces & Condiments", 2, "jars", 1, "Pantry", None, ""),
    ("Peanut butter", "Sauces & Condiments", 1, "jar", 1, "Pantry", None, ""),
    ("Honey", "Sauces & Condiments", None, "jar", 0, "Pantry", None, "how much is left?"),
    ("Tortilla chips", "Snacks", 1, "bag", 0, "Pantry", None, ""),
    ("Popcorn kernels", "Snacks", 1, "jar", 0, "Pantry", None, ""),
    ("Eggs", "Dairy & Eggs", 8, "eggs", 6, "Fridge", None, ""),
    ("Milk", "Dairy & Eggs", 0.5, "gal", 1, "Fridge", None, ""),
    ("Butter", "Dairy & Eggs", 2, "sticks", 1, "Fridge", None, ""),
    ("Shredded cheddar", "Dairy & Eggs", 1, "bag", 1, "Fridge", None, ""),
    ("Parmesan", "Dairy & Eggs", 1, "wedge", 0, "Fridge", None, ""),
    ("Greek yogurt", "Dairy & Eggs", 0, "tub", 1, "Fridge", None, ""),
    ("Onions", "Produce", 3, "onions", 2, "Counter", None, ""),
    ("Garlic", "Produce", 1, "head", 1, "Counter", None, ""),
    ("Bananas", "Produce", 4, "", 0, "Counter", None, ""),
    ("Spinach", "Produce", 1, "bag", 1, "Fridge", None, ""),
    ("Carrots", "Produce", 1, "bag", 1, "Fridge", None, ""),
    ("Bell peppers", "Produce", 0, "", 2, "Fridge", None, ""),
    ("Avocados", "Produce", 2, "", 1, "Counter", None, ""),
    ("Limes", "Produce", 1, "", 2, "Fridge", None, ""),
    ("Bread", "Bread & Tortillas", 1, "loaf", 1, "Pantry", None, ""),
    ("Flour tortillas", "Bread & Tortillas", 1, "pack", 1, "Fridge", None, ""),
    ("Frozen peas", "Frozen", 1, "bag", 1, "Freezer", None, ""),
    ("Frozen corn", "Frozen", 1, "bag", 1, "Freezer", None, ""),
    ("Veggie dumplings", "Frozen", 1, "bag", 0, "Freezer", None, ""),
    ("Ice cream", "Frozen", 1, "pint", 0, "Freezer", None, ""),
    ("Veggie chili", "Leftovers", 1, "container", 0, "Fridge", day(1), ""),
    ("Mac & cheese", "Leftovers", 1, "container", 0, "Fridge", day(6), ""),
]

# (dish, category, cuisine, status, notes, shopping list)
DISHES = [
    ("Black bean tacos", "Dinner", "Mexican", "Made before", "", "flour tortillas, 1 can black beans, onion, avocado, lime, shredded cheddar, cumin"),
    ("Spaghetti marinara", "Dinner", "Italian", "Made before", "", "spaghetti, marinara, garlic, parmesan, olive oil"),
    ("Chickpea curry", "Dinner", "Indian", "Idea", "", "1 can chickpeas, coconut milk, onion, garlic, diced tomatoes, cumin, rice"),
    ("Veggie fried rice", "Dinner", "Chinese", "Made before", "", "rice, 2 eggs, frozen peas, carrots, soy sauce, garlic, green onion"),
    ("Spinach omelette", "Breakfast", "American", "Made before", "", "3 eggs, spinach, shredded cheddar or feta, butter"),
    ("PB banana toast", "Breakfast", "American", "Made before", "", "bread, peanut butter, bananas, honey"),
    ("Stuffed peppers", "Dinner", "American", "Idea", "", "bell peppers, rice, black beans, diced tomatoes, shredded cheddar, cumin"),
    ("Popcorn", "Snack", "American", "Made before", "", "popcorn kernels, butter, salt"),
]

wb = Workbook()                                       # new empty workbook
ws = wb.active; ws.title = "Pantry"                   # first sheet = Pantry
ws.append(["Item", "Category", "Qty on Hand", "Unit", "Min to Keep", "Status", "Location", "Last Updated", "Notes"])  # header row
for r, (name, cat, qty, unit, mn, loc, upd, notes) in enumerate(ITEMS, start=2):  # one row per item
    ws.append([name, cat, qty, unit, mn, f'=IF(C{r}="","?",IF(C{r}<=0,"Out",IF(C{r}<E{r},"Low","OK")))', loc, upd, notes])
    ws.cell(r, 8).number_format = "mm/dd/yyyy"        # show dates as dates
mh = wb.create_sheet("Meal History")                  # dishes sheet
mh.append(["Dish", "Category", "Cuisine", "Status", "Notes", "Shopping List (for 2)"])
for d in DISHES:
    mh.append(list(d))
dw = wb.create_sheet("Dinner Wheel")                  # second dishes sheet (one example)
dw.append(["Dish", "Wheel Category", "Made it? (Yes/No)", "Shopping List"])
dw.append(["Pasta e fagioli", "Soups & Stews", "Yes", "penne, 1 can chickpeas, diced tomatoes, onion, garlic, carrots, parmesan"])
from openpyxl.worksheet.table import Table           # make the Pantry rows an Excel table (Export to Excel expects one)
ws.add_table(Table(displayName="Pantry", ref=f"A1:I{len(ITEMS) + 1}"))
wb.save(OUT)                                          # write the file
print("Saved", OUT)
