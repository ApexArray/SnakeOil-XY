# WIP: Script is not ready yet
"""
Run with FreeCAD's bundled interpreter.

Example:
    "C:/Program Files/FreeCAD 0.19/bin/python.exe" generate-bom.py
"""

from pathlib import Path
import re
import FreeCAD as App
import json
import FreeCADGui as Gui
from dataclasses import dataclass, field
from enum import Enum
from typing import List, Dict

# This is cross-platform now as long as the script is in the project directory
target_file = Path.home().joinpath('dev/SnakeOil-XY/CAD/v1-180-assembly.FCStd')
screenshot_out_path = Path.home().joinpath('tmp/')
# Regex pattern to match all fasteners
fastener_pattern = re.compile('.*-(Screw|Washer|HeatSet|Nut)')
# bom = {'fasteners': {}, 'other': {}}
# fasteners_bom = bom['fasteners']
# other_bom = bom['other']
type_dictionary = {}  # for debugging purposes



@dataclass
class ItemType(Enum):
    """BOM item types"""
    PRINTED = "Printed part"
    FASTENER = "Fastener (purchase)"
    OTHER = "Other (purchase)"


@dataclass
class BomItem:
    name: str  # BOM item name to be displayed
    part: App.Part  # Reference to FreeCAD part object
    item_type: ItemType  # What type of BOM entry (printed, fastener, other)
    qty: int = 0  # Quantity

    def increment(self):
        self.qty += 1

    @property
    def document(self):
        return self.part.Document

    def __repr__(self):
        return self.qty

    def __eq__(self, other):
        return self.name == other

    def __str__(self):
        return self.name


@dataclass
class BOM:
    parts: Dict[str, BomItem] = field(default_factory=dict)

    def add(self, part: App.Part):
        """Add part to BOM (auto-detects type)"""
        if fastener_pattern.match(part.Label):
            self._add_fastener(part)
        else:
            self._add_other_part_to_bom(part)

    def _add_fastener(self, part: App.Part):
        fastener = part.Label
        # Prepare fastener string to be added to BOM
        # ISO type with descriptive name
        if hasattr(part, 'type'):
            if part.type == "ISO4762":
                fastener = f"Socket head {fastener}"
            if part.type == "ISO7380-1":
                fastener = f"Button head {fastener}"
            if part.type == "ISO4026":
                fastener = f"Grub {fastener}"
            if part.type == "ISO4032":
                fastener = f"Hex {fastener}"
            if part.type == "ISO7092":
                fastener = f"Small size {fastener}"
            if part.type == "ISO7093-1":
                fastener = f"Big size {fastener}"
            if part.type == "ISO7089":
                fastener = f"Standard size {fastener}"
            if part.type == "ISO7090":
                fastener = f"Standard size {fastener}"
        # Remove numbers at end if they exist (e.g. 'M3-Washer004' becomes 'M3-Washer')
        while fastener[-1].isnumeric():
            fastener = fastener[:-1]
        # Increment fastener qty if it already exists, otherwise set fastener qty to 1
        if fastener in self.parts.keys():
            self.parts[fastener].increment()
        else:
            self.parts[fastener] = BomItem(fastener, part, item_type=ItemType.FASTENER)

    def _add_other_part_to_bom(self, part: App.Part):
        part_name = part.Label
        while part_name[-1].isnumeric() or part_name[-1] == '-':
            part_name = part_name[:-1]
        if part_name in self.parts.keys():
            self.parts[part_name].increment()
        else:
            self.parts[part_name] = BomItem(part_name, part, ItemType.OTHER)
            # get_screenshot(part.Document, part)


bom = BOM()
def get_bom_from_freecad_document(assembly: App.Document):
    print("# Getting BOM from", assembly.Name)
    for part in assembly.findObjects(Type='App::Part'):
        bom.add(part)
    # Recurse through each linked file
    for linked_file in assembly.findObjects("App::Link"):
        # print("# Getting fasteners from", linked_file.Name)
        get_bom_from_freecad_document(linked_file.LinkedObject.Document)


def get_parts_from_freecad_document(document: App.Document):
    print("# Getting fasteners from", document.Name)
    return [x for x in document.Objects if x.TypeId.startswith('Part::')]


def get_screenshot(assembly: App.Document, selected_part=None):
    """Save screenshots to out_dir.  If part is given, only that part will be shown"""
    Gui.setActiveDocument(assembly.Name)
    activeView = Gui.activeView()
    App.setActiveDocument(assembly.Name)
    # if part is given, hide everything else and activate that part
    if selected_part:
        activeView.setActiveObject('part', selected_part)
        for p in Gui.ActiveDocument.Document.Objects:
            # Hide everything except for our selected part
            p.Visibility = False
        selected_part.Visibility = True
        for p in selected_part.Parents:
            p[0].Visibility = True
        activeView.viewIsometric()
        activeView.fitAll()
        print(f"Saving image to ~/{selected_part.Label}.png")
        activeView.saveImage(str(screenshot_out_path.joinpath(selected_part.Label + '.png')), 555, 555)

    else:
        activeView.setActiveObject('part', assembly.getObject('Part'))
        activeView.viewIsometric()
        activeView.fitAll()
        print(f"Saving image to ~/{assembly.Name}.png")
        activeView.saveImage(str(screenshot_out_path.joinpath(assembly.Name+'.png')), 555, 555)


if __name__ == "__main__":
    showGUI = True

    if showGUI:
        try:
            Gui.showMainWindow()
        except:
            print("Running as macro")

    print(f"# Getting fasteners from {target_file}")
    # Get assembly object from filepath
    cad_assembly = App.open(str(target_file))
    get_bom_from_freecad_document(cad_assembly)
    final_bom = [(part.name, part.qty) for part in bom.parts]

    # Pretty print BOM dictionary
    print(json.dumps(bom, indent=4))

# with open('bom-fasteners.json', 'w') as file:
#     file.write(json.dumps(fasteners_bom, indent=4))
