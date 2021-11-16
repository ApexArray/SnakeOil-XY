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
import time

try:
    Gui.showMainWindow()
except:
    print("Running as macro")

# This is cross-platform now as long as the script is in the project directory
target_file = Path.home().joinpath('dev/SnakeOil-XY/CAD/v1-180-assembly.FCStd')
screenshot_out_path = Path.home().joinpath('tmp/')
# Regex pattern to match all fasteners
fastener_pattern = re.compile('.*-(Screw|Washer|HeatSet|Nut)')
bom = {'fasteners': {}, 'other': {}}
fasteners_bom = bom['fasteners']
other_bom = bom['other']
type_dictionary = {}  # for debugging purposes


def add_fastener_to_bom(part):
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
    if fastener in fasteners_bom.keys():
        fasteners_bom[fastener] += 1
    else:
        fasteners_bom[fastener] = 1


def add_other_part_to_bom(part):
    part_name = part.Label

    while part_name[-1].isnumeric() or part_name[-1] == '-':
        part_name = part_name[:-1]
    if part_name in other_bom.keys():
        other_bom[part_name] += 1
    else:
        other_bom[part_name] = 1
        get_screenshot(part.Document, part)


def get_bom_from_freecad_document(assembly: App.Document):
    parts = []
    # Get parts from this document
    parts += [x for x in assembly.Objects if x.TypeId.startswith('Part::')]
    get_screenshot(assembly)
    for part in parts:
        # [debugging] Add type to dictionary so we can see which parts we want to add or filter out
        type_dictionary[part.Label] = part.TypeId
        # If fastener matches, add it to BOM
        if fastener_pattern.match(part.Label):
            add_fastener_to_bom(part)
        else:
            add_other_part_to_bom(part)
    # Recurse through each linked file
    for linked_file in assembly.findObjects("App::Link"):
        print("# Getting fasteners from", linked_file.Name)
        get_bom_from_freecad_document(linked_file.LinkedObject.Document)


def get_screenshot(assembly: App.Document, selected_part=None):
    """Save screenshots to out_dir.  If part is given, only that part will be shown"""
    Gui.setActiveDocument(assembly.Name)
    App.setActiveDocument(assembly.Name)
    activeView = Gui.activeView()
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

print(f"# Getting fasteners from {target_file}")
# Get assembly object from filepath
cad_assembly = App.open(str(target_file))
get_bom_from_freecad_document(cad_assembly)

# Pretty print BOM dictionary
print(json.dumps(bom, indent=4))

with open('bom-fasteners.json', 'w') as file:
    file.write(json.dumps(fasteners_bom, indent=4))
