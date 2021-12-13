"""
Picture Drawer for Fox Socrebug

@author: Andrew Kicklighter
"""

from PIL import Image, ImageDraw, ImageFont
from sheets_functions import *
import sys


"""
Draw the score bug
"""


def draw_standings_table(conference):
    image = Image.new('RGB', (200, 500))
    image.save(conference + "_standings.png", "PNG")

    standings_data = get_standings_data(conference)
    if standings_data is not None:
        print("Found data")
        # TODO draw standings



