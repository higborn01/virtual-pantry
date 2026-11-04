"""Draws the app icon (a cartoony fridge).

Run:  python make_icon.py [website_pantry_folder]
- Always saves pantry.ico here (desktop shortcut + window icon).
- If a website folder is given, also saves the web/iPhone icons there:
  icon-180.png (iPhone home screen), icon-192.png + icon-512.png (web app), favicon.png (browser tab).
Needs Pillow (comes with Anaconda).
"""
import os                         # file paths
import sys                        # command-line arguments
from PIL import Image, ImageDraw  # Pillow: make and draw on images

HERE = os.path.dirname(os.path.abspath(__file__))  # this folder
OUTLINE = "#1f2d3d"               # same dark outline as the app
S = 256                           # draw at 256x256, then resize


def draw(full_bleed=False):
    """The fridge picture. full_bleed=True fills the whole square (iPhone rounds the corners itself)."""
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))  # blank, fully transparent square
    d = ImageDraw.Draw(img)                        # a "pen" for drawing on it
    if full_bleed:                                 # iPhone/web: sky colour edge to edge, no border
        d.rectangle((0, 0, S, S), fill="#9fd8ff")
        d.ellipse((-20, 196, 276, 290), fill="#ffffff")  # snowy ground
    else:                                          # desktop: rounded tile with a dark border
        d.rounded_rectangle((20, 20, 236, 236), radius=48, fill="#9fd8ff", outline=OUTLINE, width=10)  # icy-blue background
        d.ellipse((10, 196, 246, 280), fill="#ffffff")  # snowy ground bump at the bottom
        d.rounded_rectangle((20, 20, 236, 236), radius=48, outline=OUTLINE, width=10)  # redraw the border over the snow
    d.rounded_rectangle((74, 40, 182, 214), radius=22, fill="#ffffff", outline=OUTLINE, width=9)  # fridge body
    d.line((80, 104, 176, 104), fill=OUTLINE, width=8)                 # line between freezer and fridge doors
    d.rounded_rectangle((154, 58, 166, 92), radius=6, fill="#cfdbe7", outline=OUTLINE, width=5)   # freezer handle
    d.rounded_rectangle((154, 118, 166, 176), radius=6, fill="#cfdbe7", outline=OUTLINE, width=5)  # fridge handle
    d.rounded_rectangle((92, 46, 140, 56), radius=5, fill="#e6f5ff")   # glossy shine on the freezer door
    d.rounded_rectangle((90, 130, 118, 158), radius=7, fill="#7ed957", outline=OUTLINE, width=5)  # little green item tile
    d.rounded_rectangle((96, 166, 140, 194), radius=7, fill="#ff6b81", outline=OUTLINE, width=5)  # little red item tile
    d.rounded_rectangle((86, 210, 108, 226), radius=5, fill="#4fb3e8", outline=OUTLINE, width=5)  # left foot
    d.rounded_rectangle((148, 210, 170, 226), radius=5, fill="#4fb3e8", outline=OUTLINE, width=5)  # right foot
    return img                                     # the finished picture


ico = os.path.join(HERE, "pantry.ico")             # desktop icon file
draw().save(ico, sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])  # .ico holds several sizes
print("Saved", ico)

if len(sys.argv) > 1:                              # a website folder was given
    web = sys.argv[1]                              # where to put the web icons
    square = draw(full_bleed=True).convert("RGB")  # iPhone icons can't be see-through, so drop transparency
    for size, name in [(180, "icon-180.png"), (192, "icon-192.png"), (512, "icon-512.png")]:  # each size needed
        square.resize((size, size), Image.LANCZOS).save(os.path.join(web, name))  # smooth resize + save
    draw().resize((64, 64), Image.LANCZOS).save(os.path.join(web, "favicon.png"))  # browser tab icon (keeps rounded look)
    print("Saved web icons in", web)
