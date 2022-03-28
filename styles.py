from enum import Enum
from openpyxl.styles import Font

class Colors(Enum):
    YELLOW = "00FFFF00"
    RED = "00FF0000"
    BLUE = "000000FF"
    MAGENTA = "00FF00FF"
    GREEN = "0000FF00"
   

class FontStyles(Enum):
    ARIAL_BOLD = Font(name='Arial', bold=True, italic=False)
    ARIAL_ITALIC = Font(name='Arial',bold=False, italic=True)
    ARIAL_PLAIN = Font(name='Atial', bold=False, italic=False)