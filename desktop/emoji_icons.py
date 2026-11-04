"""Picks an emoji for each pantry item and turns it into a small colour picture for the canvas.

tkinter on Windows draws emoji in black-and-white, so we draw them in colour with
Pillow (comes with Anaconda) and hand tkinter a ready-made image instead.
To change an item's emoji, edit KEYWORDS below - first match wins.
"""
import re                                         # regular expressions, for whole-word matching

try:                                              # Pillow might not be installed...
    from PIL import Image, ImageDraw, ImageFont, ImageTk  # image tools + the bridge to tkinter
except ImportError:                               # ...in which case the app just shows no emoji
    Image = None                                  # marker: Pillow is missing

# (word to look for in the item name, emoji). Checked top to bottom; the FIRST match wins,
# so specific phrases ('hot sauce', 'peanut butter') come before general words ('carrot', 'butter').
KEYWORDS = [
    # --- specific phrases that would otherwise match the wrong thing ---
    ("instant pot", "🍲"), ("protein powder", "💪"), ("moringa", "🌿"),
    ("orange juice", "🍊"), ("lemon juice", "🍋"), ("hot sauce", "🌶"),
    ("chocolate sauce", "🍫"), ("chocolate-dipped", "🍫"), ("brownie", "🍫"),
    ("peanut butter", "🥜"), ("maple butter", "🍁"), ("rice-a-roni", "🍚"),
    ("rice vinegar", "🍶"), ("egg noodle", "🍜"), ("eggplant", "🍆"),
    ("sweet potato", "🍠"), ("pumpkin", "🎃"), ("popcorn", "🍿"),
    ("beaver nugget", "🍿"), ("ice cream", "🍨"), ("cream cheese", "🧀"),
    ("sour cream", "🥣"), ("whipping cream", "🥛"), ("creamer", "☕"),
    ("celery salt", "🧂"), ("bell pepper", "🫑"), ("black pepper", "🧂"),
    ("green onion", "🌱"), ("olive oil", "🫒"), ("vegetable oil", "🌻"),
    ("sesame oil", "🥢"), ("soy sauce", "🥢"), ("taco", "🌮"),
    ("pasta sauce", "🍅"), ("marinara", "🍅"), ("sun-dried", "🍅"),
    ("canned", "🥫"), ("cake mix", "🎂"), ("pie crust", "🥧"),
    ("garam masala", "🍛"), ("curry", "🍛"), ("italian seasoning", "🌿"),
    ("chick'n", "🌱"), ("beyond steak", "🌱"), ("stir-fry", "🥢"),
    ("fruit pearl", "🍡"), ("bouillon", "🍲"), ("instant noodle", "🍜"),
    ("tortilla chip", "🌮"), ("coconut", "🥥"), ("banana", "🍌"), ("dumpling", "🥟"), ("oat", "🌾"),
    ("yogurt", "🥣"),
    # --- baking ---
    ("extract", "🧪"), ("vanilla", "🧪"), ("yeast", "🍞"), ("flour", "🌾"),
    ("sugar", "🍬"), ("sprinkle", "🧁"), ("cornstarch", "🌽"),
    # --- dairy & eggs ---
    ("cheese", "🧀"), ("colby", "🧀"), ("mozzarella", "🧀"), ("parmesan", "🧀"),
    ("cotija", "🧀"), ("milk", "🥛"), ("butter", "🧈"), ("egg", "🥚"),
    # --- produce ---
    ("avocado", "🥑"), ("carrot", "🥕"), ("cabbage", "🥬"), ("celery", "🥬"),
    ("cilantro", "🌿"), ("parsley", "🌿"), ("lime", "🍋"), ("lemon", "🍋"),
    ("garlic", "🧄"), ("onion", "🧅"), ("potato", "🥔"), ("broccoli", "🥦"),
    ("spinach", "🥬"), ("corn", "🌽"), ("tomato", "🍅"), ("peach", "🍑"),
    ("cherr", "🍒"), ("olive", "🫒"), ("bean", "🫘"), ("split pea", "🫘"),
    # --- grains, bread, snacks ---
    ("penne", "🍝"), ("rotini", "🍝"), ("macaroni", "🍝"), ("pasta", "🍝"),
    ("noodle", "🍜"), ("rice", "🍚"), ("grits", "🌾"), ("bread", "🍞"),
    ("tortilla", "🌯"), ("potsticker", "🥟"), ("oreo", "🍪"), ("cracker", "🍘"),
    ("chips", "🍟"), ("cheetos", "🍟"), ("pistachio", "🥜"), ("peanut", "🥜"),
    ("bars", "🌾"), ("pie", "🥧"), ("tart", "🥧"),
    # --- sauces & condiments ---
    ("honey", "🍯"), ("ketchup", "🍅"), ("jam", "🍓"), ("preserve", "🍓"),
    ("mayo", "🥚"), ("salsa", "🌶"), ("sambal", "🌶"), ("sriracha", "🌶"),
    ("chipotle", "🌶"), ("jerk", "🌶"), ("vinegar", "🍶"),
    # --- spices ---
    ("chili", "🌶"), ("chile", "🌶"), ("paprika", "🌶"), ("cinnamon", "🪵"),
    ("thyme", "🌿"), ("oregano", "🌿"), ("marjoram", "🌿"), ("coriander", "🌿"),
    ("msg", "✨"), ("salt", "🧂"),
]

CATEGORY_EMOJI = {                                # used when no keyword matches
    "Produce": "🥬", "Dairy & Eggs": "🥛", "Frozen": "🧊", "Canned & Dry": "🥫",
    "Grains & Pasta": "🌾", "Baking": "🧁", "Spices": "🧂", "Sauces & Condiments": "🍶",
    "Bread & Tortillas": "🍞", "Snacks": "🍿", "Leftovers": "🍱", "Other": "📦",
}


def emoji_for(item):
    """Best emoji for an item: keyword in its name, else a can for canned goods, else its category's."""
    name = item["name"].lower()                   # compare in lower case
    for word, emoji in KEYWORDS:                  # try each keyword in order
        if re.search(r"\b" + re.escape(word), name):  # found at the START of a word? (\b = word edge, so 'egg' isn't found in 'veggie')
            return emoji                          #   use it
    if (item["unit"] or "").lower() in ("can", "cans"):  # sold in cans?
        return "🥫"                               #   tin can
    return CATEGORY_EMOJI.get(item["category"], "📦")  # category fallback (box if unknown)


_font = None    # the emoji font, loaded once on first use
_cache = {}     # (emoji, size) -> finished tkinter image, so each emoji is only drawn once


def emoji_image(emoji, size, faded=False):
    """A size x size colour picture of the emoji for tkinter (pale if faded), or None if Pillow isn't available."""
    global _font                                  # we may set the module-level _font
    if Image is None:                             # no Pillow
        return None                               #   no picture
    key = (emoji, size, faded)                    # cache key (normal and faded versions are cached separately)
    if key not in _cache:                         # not drawn yet?
        if _font is None:                         # first emoji ever: load Windows' colour emoji font
            _font = ImageFont.truetype("seguiemj.ttf", 96)  # Segoe UI Emoji, drawn big for quality
        big = Image.new("RGBA", (160, 160), (0, 0, 0, 0))  # transparent scratch canvas
        ImageDraw.Draw(big).text((16, 16), emoji, font=_font, embedded_color=True)  # draw the emoji in colour
        box = big.getbbox()                       # tight box around the non-transparent pixels
        if box is None:                           # font had no picture for this emoji (blank result)
            _cache[key] = None                    #   remember that, show nothing
            return None
        big = big.crop(box)                       # cut away the empty border
        big.thumbnail((size, size), Image.LANCZOS)  # shrink to fit size x size (LANCZOS = smooth, sharp shrinking)
        small = Image.new("RGBA", (size, size), (0, 0, 0, 0))  # final transparent square
        small.paste(big, ((size - big.width) // 2, (size - big.height) // 2))  # centre the emoji in it
        if faded:                                 # pale version for tiles that don't match the search
            alpha = small.getchannel("A")         # the transparency layer (0 = invisible, 255 = solid)
            small.putalpha(alpha.point(lambda a: a * 30 // 100))  # make every pixel 30% as solid
        _cache[key] = ImageTk.PhotoImage(small)   # convert to a tkinter image (kept in the cache so it isn't deleted)
    return _cache[key]                            # the finished picture
