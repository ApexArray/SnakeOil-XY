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

Gui.showMainWindow()

# This is cross-platform now as long as the script is in the project directory
target_file = Path(__file__).parent.joinpath('CAD/v1-180-assembly.FCStd')
bom_out_dir = Path(__file__).parent.joinpath('bom-out')
# Regex pattern to match all fasteners
fastener_pattern = re.compile('.*-(Screw|Washer|HeatSet|Nut)')
bom = {'fasteners': {}, 'other': {}, 'printed': {}}
fasteners_bom = bom['fasteners']
other_bom = bom['other']
printed_bom = bom['printed']
component_bom = {}  # BOM for each individual component
type_dictionary = {}  # for debugging purposes
printed_parts_colors = [
    (0.3333333432674408, 1.0, 1.0, 0.0),  # Teal
    (0.6666666865348816, 0.6666666865348816, 1.0, 0.0)  # Blue
]


def add_fastener_to_bom(part, component=None):
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

    if component:
        if component not in component_bom.keys():
            component_bom[component] = {'fasteners': {}, 'other': {}, 'printed': {}}
        if fastener in component_bom[component]['fasteners'].keys():
            component_bom[component]['fasteners'][fastener] += 1
        else:
            component_bom[component]['fasteners'][fastener] = 1


def add_other_part_to_bom(part, component):
    part_name = part.Label

    while part_name[-1].isnumeric() or part_name[-1] == '-':
        part_name = part_name[:-1]
    if part_name in other_bom.keys():
        other_bom[part_name] += 1
    else:
        other_bom[part_name] = 1
    if component:
        if component not in component_bom.keys():
            component_bom[component] = {'fasteners': {}, 'other': {}, 'printed': {}}
        if part_name in component_bom[component]['other'].keys():
            component_bom[component]['other'][part_name] += 1
        else:
            component_bom[component]['other'][part_name] = 1


def add_printed_part_to_bom(part, component):
    part_name = part.Label
    if part_name in other_bom.keys():
        other_bom[part_name] += 1
    else:
        other_bom[part_name] = 1
    if component:
        if component not in component_bom.keys():
            component_bom[component] = {'fasteners': {}, 'other': {}, 'printed': {}}
        if part_name in component_bom[component]['printed'].keys():
            component_bom[component]['printed'][part_name] += 1
        else:
            component_bom[component]['printed'][part_name] = 1


def get_bom_from_freecad_document(assembly: App.Document):
    # parts = []
    # Get parts from this document
    # parts +=
    for part in [x for x in assembly.Objects if x.TypeId.startswith('Part::')]:
        # [debugging] Add type to dictionary so we can see which parts we want to add or filter out
        type_dictionary[part.Label] = part.TypeId
        # If fastener matches, add it to BOM
        if fastener_pattern.match(part.Label):
            add_fastener_to_bom(part, assembly.Label)
        elif part.ViewObject.ShapeColor in printed_parts_colors:
            add_printed_part_to_bom(part, assembly.Label)
        else:
            add_other_part_to_bom(part, assembly.Label)

    if assembly.Label in component_bom.keys():
        with open(bom_out_dir.joinpath(f'bom-{assembly.Label}.json'), 'w') as f:
            f.write(json.dumps(component_bom[assembly.Label], indent=4))
    # Recurse through each linked file
    for linked_file in assembly.findObjects("App::Link"):
        print("# Getting fasteners from", linked_file.Name)
        get_bom_from_freecad_document(linked_file.LinkedObject.Document)


print(f"# Getting fasteners from {target_file}")
# Get assembly object from filepath
cad_assembly = App.open(str(target_file))
get_bom_from_freecad_document(cad_assembly)

# Pretty print BOM dictionary
print(json.dumps(bom, indent=4))

with open('bom-fasteners.json', 'w') as file:
    file.write(json.dumps(component_bom, indent=4))

