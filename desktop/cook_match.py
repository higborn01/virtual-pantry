"""'What can I cook?': matches each dish's ingredient keywords against what's in stock.

Tweak the three lists below to improve matching.
"""
import re                          # regular expressions, to split names into words

from import_excel import singular  # same plural -> singular rule the import uses

# Ingredients everyone has - never counted as missing.
STAPLES = {"salt", "pepper", "black pepper", "water", "oil", "olive oil", "vegetable oil",
           "salt and pepper", "cooking spray", "ice"}

# Recipe keyword -> other things in the pantry that count as having it.
SYNONYMS = {
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

# Words too general to match on their own (a recipe's 'hot sauce' shouldn't match any '... sauce').
GENERIC = {"bean", "sauce", "oil", "powder", "paste", "seasoning", "mix", "juice", "pepper",
           "vinegar", "cream", "extract", "spice", "cheese", "flake", "seed", "leaf", "cup", "bar"}


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
        self.profiles = []                               # one (all words, core words) pair per in-stock item
        for item in items:                               # every item
            if not in_stock(item):                       # skip Out items
                continue
            name = re.sub(r"\([^)]*\)", " ", item["name"].lower())  # name without brackets
            alts = re.split(r"/|\bor\b", name)           # 'Tortillas / wraps' -> two alternatives
            cores = {words(a)[-1] for a in alts if words(a)}  # last word of each alternative: 'tortilla', 'wrap'
            self.profiles.append((set(words(name)), cores))   # remember both

    def has(self, keyword):
        """True if some in-stock item matches the keyword (or one of its synonyms)."""
        for phrase in [keyword] + SYNONYMS.get(keyword, []):  # the keyword itself, then its synonyms
            kw = words(phrase)                           # its words
            if not kw:                                   # nothing left (e.g. only symbols)
                continue
            head = kw[-1]                                # main noun = last word ('crusty bread' -> 'bread')
            for all_words, cores in self.profiles:       # check every in-stock item
                if set(kw) <= all_words:                 # rule 1: every keyword word is in the item name
                    return True
                if head not in GENERIC and head in cores:  # rule 2: same main noun ('flour tortilla' ~ 'Tortillas')
                    return True
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
