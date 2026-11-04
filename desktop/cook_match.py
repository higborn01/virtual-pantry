"""'What can I cook?': matches each dish's ingredient keywords against what's in stock.

Tweak the three lists below to improve matching.
"""
import re                          # regular expressions, to split names into words

from import_excel import singular  # same plural -> singular rule the import uses

# Ingredients everyone has - never counted as missing.
STAPLES = {"salt", "pepper", "black pepper", "water", "oil", "olive oil", "vegetable oil",
           "salt and pepper", "cooking spray", "ice"}

# Describing words that don't change WHAT the food is ('shredded mozzarella' is still mozzarella).
SOFT = set("""fresh frozen dried canned chopped diced sliced minced shredded grated ground whole baby large small
    extra unsalted salted plain organic raw cooked day old jumbo mini of and the a pure fancy crushed roasted red
    yellow white brown sharp mild heavy firm wide fried boiled soft hard scrambled poached caramelized toasted
    steamed crusty sandwich long grain basmati jasmine regular light opened unopened labeled russet ceylon elbow""".split())

# Brand / store words found in item names (ignored when matching).
BRANDS = set("""gv great value barilla trader joe newman own heinz king arthur fleischmann mccormick cacique marukan
    lee kum kee one dakin farm walkerswood san marco tai pei gardein nature valley duncan hine dolce flav
    savannah bee garden life golden publix better than""".split())

# Container words: ignored in item names ('Unsalted butter sticks' is butter) and in recipe keywords
# (except 'stick' and 'cup' in recipes - 'mozzarella sticks' is a different food from mozzarella!).
PACKAGING = {"jar", "bag", "can", "box", "bottle", "container", "pack", "pkg", "carton", "tub", "block",
             "shaker", "pint", "gal", "bunch", "stick", "cup"}
KEEP_IN_RECIPES = {"stick", "cup"}

# Recipe keyword -> other things in the pantry that count as having it.
SYNONYMS = {
    "hot sauce": ["sriracha", "buffalo hot sauce", "carrot hot sauce", "honey hot sauce"],
    "salsa": ["salsa verde"], "vinegar": ["balsamic vinegar", "rice vinegar"],
    "chili paste": ["sambal oelek"], "sambal": ["sambal oelek"],
    "curry roux": ["curry"], "curry block": ["curry"], "japanese curry": ["curry"],
    "vegetable broth": ["bouillon"], "veg broth": ["bouillon"], "broth": ["bouillon"],
    "vegetable stock": ["bouillon"], "stock": ["bouillon"],
    "scallion": ["green onion"],
    "cheddar": ["cheese"], "sharp cheddar": ["cheese"], "shredded cheddar": ["cheese", "colby"],
    "shredded cheese": ["colby", "mozzarella", "cheese"], "mexican blend": ["colby"],
    "pasta": ["penne", "rotini", "macaroni"], "elbow pasta": ["macaroni"],
    "spaghetti": ["pasta", "penne", "rotini"], "spaghetti noodle": ["pasta", "penne", "rotini"],
    "marinara": ["pasta sauce"], "tomato sauce": ["marinara", "pasta sauce"],
    "crusty bread": ["bread"], "sourdough": ["bread"], "bun": ["bread"],
    "heavy cream": ["whipping cream"], "cream": ["whipping cream"],
    "lemon": ["lemon juice"], "lime juice": ["lime"],
}

def words(text):
    """'Black Beans (canned)' -> ['black', 'bean']: lower-case words, singular, brackets dropped."""
    text = re.sub(r"\([^)]*\)", " ", text.lower())   # remove (bracketed) bits
    return [singular(w) for w in re.findall(r"[^\W\d_]+", text)]  # letters-only words (handles é etc.), each made singular


def in_stock(item):
    """Counts as 'have it': quantity unknown, or more than 0 (so OK, Low or ?)."""
    return item["qty"] is None or item["qty"] > 0


class Matcher:
    """Answers 'do we have <keyword>?' using the in-stock items."""

    def __init__(self, items):
        self.foods = []                                  # one (main word, other words) pair per food in stock
        for item in items:                               # every item
            if not in_stock(item):                       # skip Out items
                continue
            name = re.sub(r"\([^)]*\)", " ", item["name"].lower())  # name without (brackets)
            for alt in re.split(r"/|\bor\b", name):      # 'Tortillas / wraps' -> 'tortillas', 'wraps' (each is a food)
                w = [x for x in words(alt) if x not in PACKAGING and x not in BRANDS and len(x) > 1]  # drop jar/GV/etc.
                if w:                                    # anything left?
                    self.foods.append((w[-1], set(w[:-1])))  # main word = last word; the rest describe it

    def has(self, keyword):
        """True if some in-stock food IS this ingredient (or one of its synonyms)."""
        for phrase in [keyword] + SYNONYMS.get(keyword, []):  # the keyword itself, then its synonyms
            kw = [x for x in words(phrase) if x not in PACKAGING - KEEP_IN_RECIPES]  # recipe words minus containers
            if not kw:                                   # nothing left
                continue
            head, extra = kw[-1], set(kw[:-1])           # main word + describing words, e.g. 'mozzarella' + {'fresh'}
            for food_head, food_extra in self.foods:     # every food in stock
                if (food_head == head                    # same main word ('cheese' = 'cheese')...
                        and food_extra <= extra | SOFT   # ...the pantry item's extra words are in the recipe or harmless
                        and extra <= food_extra | SOFT):  # ...and the recipe's extra words are in the item or harmless
                    return True                          #   (so 'peanut butter' is NOT 'butter', 'coconut milk' is NOT 'milk')
        return False                                     # nothing matched


def rank(dishes, groups, items):
    """Score every dish. `groups` = {dish_id: {alt_group: [(raw, keyword), ...]}}.
    Returns a list of dicts, best first: dish, have, total, pct, missing (raw texts)."""
    m = Matcher(items)                                   # build the stock lookup once
    results = []                                         # one entry per dish
    for dish in dishes:                                  # every dish
        have, missing = 0, []                            # counters for this dish
        for options in groups.get(dish["id"], {}).values():  # each ingredient (with its alternatives)
            keywords = [kw for _, kw in options]         # e.g. ['feta', 'mozzarella']
            if any(kw in STAPLES for kw in keywords):    # a staple -> don't count it at all
                continue
            if any(m.has(kw) for kw in keywords):        # have any of the alternatives?
                have += 1
            else:
                missing.append(options[0][0])            # missing: remember its original text, e.g. 'feta or mozzarella'
        total = have + len(missing)                      # ingredients that count
        if total:                                        # skip dishes with no (non-staple) ingredients
            results.append(dict(dish=dish, have=have, total=total, pct=round(100 * have / total), missing=missing))
    results.sort(key=lambda r: (-r["pct"], len(r["missing"]), r["dish"]["name"].lower()))  # best %, fewest missing, A-Z
    return results
