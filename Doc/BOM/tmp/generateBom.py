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
from typing import List
import FreeCAD as App
import FreeCADGui as Gui
import json
import logging
from dataclasses import dataclass, field
from enum import Enum

LOGGER = logging.getLogger(__name__)

# Open GUI if running from console, otherwise we know we are running from a macro
if hasattr(Gui, 'showMainWindow'):
    Gui.showMainWindow()
else:
    print("Running as macro")

# Use SNAKEOIL_PROJECT_PATH environment variable if exists, else default to @Chip's directory
# SNAKEOIL_PROJECT_PATH = os.getenv('SNAKEOIL_PROJECT_PATH', '/home/chip/Data/Code/SnakeOil-XY/')
BASE_PATH = Path(os.path.dirname(__file__))
SNAKEOIL_PROJECT_PATH = str(BASE_PATH.parent.parent.parent)
target_file = Path(SNAKEOIL_PROJECT_PATH).joinpath('CAD/v1-180-assembly.FCStd')
bom_out_dir = Path(SNAKEOIL_PROJECT_PATH).joinpath('Doc/BOM/tmp')
# Regex pattern to match all fasteners
fastener_pattern = re.compile('.*-(Screw|Washer|HeatSet|Nut)')
# If a shape color in this list, the object will be treated as a printed part
printed_parts_colors = [
    (0.3333333432674408, 1.0, 1.0, 0.0),  # Teal
    (0.6666666865348816, 0.6666666865348816, 1.0, 0.0),  # Blue
]

# Quick references to BOM part types.  Also provides type hinting in the BomPart dataclass
PRINTED_MAIN = "printed (main)"
PRINTED_ACCENT = "printed (accent)"
FASTENER = "fastener"
OTHER = "other"
BomItemType = Enum('BomItemType', [PRINTED_MAIN, PRINTED_ACCENT, FASTENER, OTHER])


def get_new_bom():
    """Use this factory to get a new empty BOM dict"""
    _bom = {}
    for partType in BomItemType:
        _bom[partType.name] = {}
    return _bom

def get_stl_files():
    dir = Path(SNAKEOIL_PROJECT_PATH)
    stl_files = [x for x in dir.glob("**/*.stl")]
    LOGGER.info(f"# Found {len(stl_files)} stl files")
    return stl_files

# Create new BOM dictionaries
bom = get_new_bom()
detail_bom = get_new_bom()
# Quick references
fasteners_bom = bom[FASTENER]
printed_bom = bom[PRINTED_MAIN]
printed_accent_bom = bom[PRINTED_ACCENT]
other_bom = bom[OTHER]

@dataclass
class PrintedPart:
    part: App.Part
    parent: str = field(init=False, default='')
    Label: str = field(init=False)

    def __post_init__(self):
        try:
            self.Label = self.part.Label
            self.parent = self.part.Parents[0][0].Label
            if self.parent == "snakeoilxy-180":
                self.parent = self.part.Parents[1][0].Label
        except:
            pass

    def __str__(self) -> str:
        return f"{self.part.Label} [{self.parent}]"
    
    def __repr__(self) -> str:
        return self.__str__()

@dataclass
class BomItem:
    part: App.Part  # Reference to FreeCAD part object
    type: BomItemType  # What type of BOM entry (printed, fastener, other)
    name: str = field(init=False)  # We'll get the name from the part.label in the __post_init__ function

    def __post_init__(self):
        self.name = self.part.Label
        # Remove numbers at end if they exist (e.g. 'M3-Washer004' becomes 'M3-Washer')
        while self.name[-1].isnumeric():
            self.name = self.name[:-1]
        # Add descriptive fastener names
        if self.type == FASTENER:
            if hasattr(self.part, 'type'):
                if self.part.type == "ISO4762":
                    self.name = f"Socket head {self.name}"
                if self.part.type == "ISO7380-1":
                    self.name = f"Button head {self.name}"
                if self.part.type == "ISO4026":
                    self.name = f"Grub {self.name}"
                if self.part.type == "ISO4032":
                    self.name = f"Hex {self.name}"
                if self.part.type == "ISO7092":
                    self.name = f"Small size {self.name}"
                if self.part.type == "ISO7093-1":
                    self.name = f"Big size {self.name}"
                if self.part.type == "ISO7089":
                    self.name = f"Standard size {self.name}"
                if self.part.type == "ISO7090":
                    self.name = f"Standard size {self.name}"


def _add_to_main_bom(bomItem: BomItem):
    targetBom = bom[bomItem.type]
    if bomItem.name in targetBom.keys():
        targetBom[bomItem.name] += 1
    else:
        targetBom[bomItem.name] = 1


def _add_to_detailed_bom(bomItem: BomItem):
    parentPartName = bomItem.part.Document.Label
    if parentPartName not in detail_bom.keys():
        detail_bom[parentPartName] = get_new_bom()
    targetDetailBom = detail_bom[parentPartName][bomItem.type]
    if bomItem.name in targetDetailBom.keys():
        targetDetailBom[bomItem.name] += 1
    else:
        targetDetailBom[bomItem.name] = 1

def check_part_color(part: App.Part):
    if part.ViewObject.ShapeColor == (0.3333333432674408, 1.0, 1.0, 0.0):  # Teal
        return PRINTED_MAIN
    elif part.ViewObject.ShapeColor == (0.6666666865348816, 0.6666666865348816, 1.0, 0.0):  # Blue
        return PRINTED_ACCENT
    else:
        return None

def add_to_bom(part: App.Part):
    # Sort parts by type
    if fastener_pattern.match(part.Label):
        bomItem = BomItem(part, FASTENER)
    else:
        part_color = check_part_color(part)
        if part_color is not None:
            bomItem = BomItem(part, part_color)       
        else:
            bomItem = BomItem(part, OTHER)
    _add_to_main_bom(bomItem)
    _add_to_detailed_bom(bomItem)

def read_printed_parts_from_freecad_document(assembly: App.Document) -> list[App.Part]:
    LOGGER.debug("# Getting parts from", assembly.Label)
    freecad_printed_parts = [x for x in assembly.Objects if x.TypeId.startswith('Part::')]
    # Recurse through each linked file
    for linked_file in assembly.findObjects("App::Link"):
        LOGGER.debug("# Getting linked parts from", linked_file.Name)
        freecad_printed_parts += read_printed_parts_from_freecad_document(linked_file.LinkedObject.Document)
    return freecad_printed_parts


def get_bom_from_freecad_document(printed_parts: List[App.Part]):
    for part in printed_parts:
        # if part.Label in ["knob", "wago-mount", "wago-mounter", 'nut', 'z-belt-mounter-clamp-nut']:
        #     LOGGER.warning(PrintedPart(part))
        add_to_bom(part)


def addCustomfFastener(fastenerName, count):
    if fastenerName in fasteners_bom.keys():
        fasteners_bom[fastenerName] += count
    else:
        fasteners_bom[fastenerName] = count


def sort_dictionary_recursive(dictionary):
    sortedDict = {}
    for i in sorted(dictionary):
        val = dictionary[i]
        if type(val) is dict:
            val = sort_dictionary_recursive(val)
        sortedDict[i] = val
    return sortedDict


def write_bom_to_file(target_file_name, bomContent):
    # sort dict
    sortedDict = sort_dictionary_recursive(bomContent)
    filePath = bom_out_dir.joinpath(target_file_name)
    LOGGER.info(f"# Writing to {target_file_name}")
    # Sort dictionary alphabetically by key
    with open(filePath, 'w') as bom_file:
        bom_file.write(json.dumps(sortedDict, indent=2))

global parts_dict
parts_dict = None

def load_printed_parts_from_file(filename='bom-printed-parts.json'):
    """Returns tuple main_colors, accent_colors from a json file"""
    global parts_dict
    if parts_dict is None:
        with open(filename) as file:
            parts_dict = json.load(file)
    main_colors = parts_dict['printed (main color)']
    accent_colors = parts_dict['printed (accent color)']
    return main_colors, accent_colors

def get_part_color_from_filename(file: Path, printed_parts: List[App.Part]):
    """Check if part_name is a main or accent color. Returns 'main', 'accent' or None"""
    file_name = file.name
    parts_main_color = [part for part in printed_parts if check_part_color(part) == PRINTED_MAIN]
    parts_accent_color = [part for part in printed_parts if check_part_color(part) == PRINTED_ACCENT]
    main_results = [PrintedPart(part) for part in parts_main_color if part.Label in file_name]
    accent_results = [PrintedPart(part) for part in parts_accent_color if part.Label in file_name]
    main_count = len(main_results)
    accent_count = len(accent_results)
    total_count = main_count + accent_count
    if total_count != 1:
        # print(f"# Found {total_count} results for {file_name}")
        if total_count > 1 and total_count not in [main_count, accent_count]:
            LOGGER.warning(f"# {file.relative_to(SNAKEOIL_PROJECT_PATH)} matches {total_count} part_names")
            main_list = '\n'.join([f'   - {part}' for part in main_results])
            accent_list = '\n'.join([f'   - {part}' for part in accent_results])
            print(f"  main colors:\n{main_list}")
            print(f"  accent colors\n{accent_list}\n")
    else:
        if main_count == 1:
            return "main"
        elif accent_count == 1:
            return "accent"
        else:
            raise ValueError(f"total count is {total_count}, main and accent count != 1")

def get_filename_color_report(printed_parts: List[App.Part]):
    """return dictionary of filename['main'|'accent'|None]"""
    stl_files = get_stl_files()
    # main_colors, accent_colors = load_printed_parts_from_file()
    file_results = {part_name: get_part_color_from_filename(part_name, printed_parts) for part_name in stl_files}
    # print(file_results)
    total_main_parts = len([file for (file, result)  in file_results.items() if result == 'main'])
    total_accent_parts = len([file for (file, result)  in file_results.items() if result == 'accent'])
    total_missing_parts = len([file for (file, result)  in file_results.items() if result is None])
    print(f"# Total main parts: {total_main_parts}")
    print(f"# Total accent parts: {total_accent_parts}")
    print(f"# Total missing parts: {total_missing_parts}")
    return file_results

if __name__ == '__main__':

    # if len(sys.argv) > 1:
    #     cmd = sys.argv[1]
    #     if cmd == 'generate_new_bom':
    LOGGER.info(f"# Getting BOM from {target_file}")
    # Get assembly object from filepath
    cad_assembly = App.open(str(target_file))
    printed_parts = read_printed_parts_from_freecad_document(cad_assembly)
    get_bom_from_freecad_document(printed_parts)

    # add custom Fasteners that not in the assembly
    # spring washer for rails mount
    addCustomfFastener("Spring washer M3", 60)
    # bolts for 3030 extrusion rails mount
    addCustomfFastener("Socket head M3x10-Screw", 50)
    # bolts for 1515 extrusion rail mount
    addCustomfFastener("Socket head M3x8-Screw", 10)
    # T-nut for 1515 gantry and bed
    addCustomfFastener("Square M3-Nut", 30)
    # count 3030 M6 T-nut = M6 bolt
    m6NutCount = 0
    for fastenersName in fasteners_bom.keys():
        if "Screw" in fastenersName and "M6" in fastenersName:
            m6NutCount += fasteners_bom[fastenersName]
    addCustomfFastener("3030 M6-T-nut", m6NutCount)
    # 3030 M3 t-nut (50 for rails, 10 for other add-ons)
    addCustomfFastener("3030 M3-T-nut", 60)
    # 3030 M5 t-nut for z motor mount and others
    addCustomfFastener("3030 M5-T-nut", 10)

    # Write to files
    write_bom_to_file('bom-all.json', bom)
    write_bom_to_file('bom-fasteners.json', fasteners_bom)
    write_bom_to_file('bom-printed-parts.json', {'printed (main color)': printed_bom, 'printed (accent color)': printed_accent_bom})
    write_bom_to_file('bom-detail.json', detail_bom)
    write_bom_to_file('bom-other.json', other_bom)

    print("# completed!")

    get_filename_color_report(printed_parts)