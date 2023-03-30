"""
Run with FreeCAD's bundled interpreter, or as a FreeCAD macro
"""
import os
if os.name == 'nt':
    FREECADPATH = os.getenv('FREECADPATH', 'C:/Program Files/FreeCAD 0.20/bin')
else:
    FREECADPATH = os.getenv('FREECADPATH', '/usr/lib/freecad-python3/lib/')
import sys
sys.path.append(FREECADPATH)
from pathlib import Path
import re
import sys
from typing import Dict, List, Union
import FreeCAD as App  # type: ignore
import FreeCADGui as Gui  # type: ignore