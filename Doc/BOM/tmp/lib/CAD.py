"""
Run with FreeCAD's bundled interpreter, or as a FreeCAD macro
"""
import os
FREECADPATH = os.getenv('FREECADPATH', '/usr/lib/freecad-python3/lib/')
import sys
sys.path.append(FREECADPATH)
from pathlib import Path
import re
import sys
from typing import List, Union
import FreeCAD as App
import FreeCADGui as Gui
import json
import logging
from dataclasses import InitVar, dataclass, field
from enum import Enum
import shelve

# Quick references to BOM part types.  Also provides type hinting in the BomPart dataclass
PRINTED_MAIN = "main"
PRINTED_ACCENT = "accent"
PRINTED_MISSING = "missing"
PRINTED_UNKNOWN_COLOR = "unknown"
PRINTED_CONFLICTING_COLORS = "conflicting"
FASTENER = "fastener"
OTHER = "other"

logging.basicConfig(
    filename='generateBom.log', filemode='a', format='%(levelname)s: %(message)s', level=logging.INFO
    )
LOGGER = logging.getLogger()

BomItemType = Enum('BomItemType', [PRINTED_MAIN, PRINTED_ACCENT, FASTENER, OTHER])
fastener_pattern = re.compile('.*-(Screw|Washer|HeatSet|Nut)')

def get_printed_part_color(part: App.Part):
    """Checks if CAD object is a printed part, as determined by its color (teal=main, blue=accent)

    Args:
        part (App.Part): FreeCAD object to check

    Returns:
        str: friendly name of color (ex: main, accent)
        None: if not a known color for printed parts
    """
    if part.ViewObject.ShapeColor == (0.3333333432674408, 1.0, 1.0, 0.0):  # Teal
        return PRINTED_MAIN
    elif part.ViewObject.ShapeColor == (0.6666666865348816, 0.6666666865348816, 1.0, 0.0):  # Blue
        return PRINTED_ACCENT
    else:
        return None

@dataclass
class BomItem:
    part: InitVar[App.Part]
    type: BomItemType = field(init=False)  # What type of BOM entry (printed, fastener, other)
    name: str = field(init=False)  # We'll get the name from the part.label in the __post_init__ function
    parent: str = field(init=False, default='')
    color_category: str = field(init=False)
    raw_color: tuple = field(init=False)

    def __post_init__(self, part):
        self.name = part.Label
        self.raw_color = part.ViewObject.ShapeColor
        # Remove numbers at end if they exist (e.g. 'M3-Washer004' becomes 'M3-Washer')
        while self.name[-1].isnumeric():
            self.name = self.name[:-1]
        if fastener_pattern.match(part.Label):
            self.type = FASTENER
        else:
            color_category = get_printed_part_color(part)
            if color_category is not None:
                self.type = color_category
            else:
                self.type = OTHER
        # Add descriptive fastener names
        if self.type == FASTENER:
            if hasattr(part, 'type'):
                if part.type == "ISO4762":
                    self.name = f"Socket head {self.name}"
                if part.type == "ISO7380-1":
                    self.name = f"Button head {self.name}"
                if part.type == "ISO4026":
                    self.name = f"Grub {self.name}"
                if part.type == "ISO4032":
                    self.name = f"Hex {self.name}"
                if part.type == "ISO7092":
                    self.name = f"Small size {self.name}"
                if part.type == "ISO7093-1":
                    self.name = f"Big size {self.name}"
                if part.type == "ISO7089":
                    self.name = f"Standard size {self.name}"
                if part.type == "ISO7090":
                    self.name = f"Standard size {self.name}"
        # Try to add parent object
        try:
            self.parent = part.Parents[0][0].Label
            if self.parent == "snakeoilxy-180":
                self.parent = part.Parents[1][0].Label
        except:
            pass
    
    def __str__(self) -> str:
        return f"{self.parent}/{self.name}"
    
    def __repr__(self) -> str:
        return self.__str__()
    
def get_cad_objects_from_cache() -> List[BomItem]:
    with shelve.open('cad_cache') as db:
        return db['freecad_objects']
    
def write_cad_objects_to_cache(cad_objects: List[BomItem]):
    LOGGER.info("Writing cad objects to cache")
    with shelve.open('cad_cache') as db:
        db['freecad_objects'] = cad_objects

def get_cad_objects_from_freecad(assembly: App.Document) -> List[BomItem]:
    freecad_objects = None
    # Try to read from cache to avoid length process of reading CAD file
    LOGGER.debug("# Getting parts from", assembly.Label)
    freecad_objects = [BomItem(x) for x in assembly.Objects if x.TypeId.startswith('Part::')]
    # Recurse through each linked file
    for linked_file in assembly.findObjects("App::Link"):
        LOGGER.debug("# Getting linked parts from", linked_file.Name)
        freecad_objects += get_cad_objects_from_freecad(linked_file.LinkedObject.Document)
    return freecad_objects
