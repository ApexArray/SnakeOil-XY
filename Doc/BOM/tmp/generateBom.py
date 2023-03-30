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
import json
import logging
from enum import Enum
from lib.CAD import (
    BomItem, PRINTED_MAIN, PRINTED_ACCENT, PRINTED_MISSING, PRINTED_UNKNOWN_COLOR, PRINTED_CONFLICTING_COLORS, 
    FASTENER, OTHER, get_cad_objects_from_freecad, get_cad_objects_from_cache, get_filename_color_results, write_cad_objects_to_cache
    )

logging.basicConfig(
    filename='generateBom.log', filemode='w', format='%(levelname)s: %(message)s', level=logging.INFO
    )
LOGGER = logging.getLogger()

# Use SNAKEOIL_PROJECT_PATH environment variable if exists, else default to @Chip's directory
# SNAKEOIL_PROJECT_PATH = os.getenv('SNAKEOIL_PROJECT_PATH', '/home/chip/Data/Code/SnakeOil-XY/')
BASE_PATH = Path(os.path.dirname(__file__))
SNAKEOIL_PROJECT_PATH = str(BASE_PATH.parent.parent.parent)
target_file = Path(SNAKEOIL_PROJECT_PATH).joinpath('CAD/v1-180-assembly.FCStd')
bom_out_dir = Path(SNAKEOIL_PROJECT_PATH).joinpath('Doc/BOM/tmp')
STL_PATH = (
    Path(SNAKEOIL_PROJECT_PATH) / 'BETA3_Standard_Release_STL' / 'STLs'
    ).relative_to(SNAKEOIL_PROJECT_PATH)
EXCLUDE_DIRS = [
    STL_PATH / "Add-on",
    STL_PATH / "Tools",
]
EXCLUDE_STRINGS = [
    "OPTIONAL"
]
# Regex pattern to match all fasteners
fastener_pattern = re.compile('.*-(Screw|Washer|HeatSet|Nut)')

# STL_GLOB = "**/*.stl"
STL_GLOB = f"{STL_PATH}/**/*.stl"

BomItemType = Enum('BomItemType', [PRINTED_MAIN, PRINTED_ACCENT, FASTENER, OTHER])

def get_new_bom():
    """Use this factory to get a new empty BOM dict"""
    _bom = {}
    for partType in BomItemType:
        _bom[partType.name] = {}
    return _bom

def filter_stl_files(stl_files: List[Path]):
    filtered_stl_files = []
    for file in stl_files:
        keep = True
        for excluded_fp in EXCLUDE_DIRS:
            if str(file).startswith(str(excluded_fp)):
                keep = False
        for excluded_str in EXCLUDE_STRINGS:
            if excluded_str.lower() in str(file).lower():
                keep = False
        if keep == True:
            filtered_stl_files.append(file)
    return filtered_stl_files

def get_stl_files() -> List[Path]:
    dir = Path(SNAKEOIL_PROJECT_PATH)
    stl_files = [x.relative_to(SNAKEOIL_PROJECT_PATH) for x in dir.glob(STL_GLOB)]
    stl_files = filter_stl_files(stl_files)
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

def _add_to_main_bom(bomItem: BomItem):
    targetBom = bom[bomItem.type]
    if bomItem.name in targetBom.keys():
        targetBom[bomItem.name] += 1
    else:
        targetBom[bomItem.name] = 1

def _add_to_detailed_bom(bomItem: BomItem):
    documentName = bomItem.document
    if documentName not in detail_bom.keys():
        detail_bom[documentName] = get_new_bom()
    targetDetailBom = detail_bom[documentName][bomItem.type]
    if bomItem.name in targetDetailBom.keys():
        targetDetailBom[bomItem.name] += 1
    else:
        targetDetailBom[bomItem.name] = 1

def get_bom_from_freecad_document(bom_items: List[BomItem]):
    for bomItem in bom_items:
        _add_to_main_bom(bomItem)
        _add_to_detailed_bom(bomItem)

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

def add_fasteners():
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

def write_bom_files():
    # Write to files
    write_bom_to_file('bom-all.json', bom)
    write_bom_to_file('bom-fasteners.json', fasteners_bom)
    write_bom_to_file('bom-printed-parts.json', {
        'printed (main color)': printed_bom, 
        'printed (accent color)': printed_accent_bom
        })
    write_bom_to_file('bom-detail.json', detail_bom)
    write_bom_to_file('bom-other.json', other_bom)

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

def write_filename_reports(filename_results: Dict[str, List[Union[Path, str]]]):
    for category in [
        PRINTED_MAIN, PRINTED_ACCENT, PRINTED_MISSING, PRINTED_UNKNOWN_COLOR, PRINTED_CONFLICTING_COLORS
        ]:
        with open(BASE_PATH / f'color-results-{category}.txt', 'w') as file:
            results: list[Path] = filename_results[category]
            if issubclass(type(results[0]), Path):
                results: list[str] = [x.as_posix() for x in results]
            sorted_results = sorted(results)
            file.write('\n'.join(sorted_results))

if __name__ == '__main__':
    LOGGER.info(f"# Getting BOM from {target_file}")
    # Get assembly object from filepath
    try:
        cad_parts = get_cad_objects_from_cache()
    except KeyError:
        LOGGER.info("# No cached freecad_objects found. Reading from CAD file")
        # Open GUI if running from console, otherwise we know we are running from a macro
        if hasattr(Gui, 'showMainWindow'):
            Gui.showMainWindow()
        else:
            print("Running as macro")
        cad_assembly = App.open(str(target_file))
        cad_parts = get_cad_objects_from_freecad(cad_assembly)
        write_cad_objects_to_cache(cad_parts)
    get_bom_from_freecad_document(cad_parts)
    # Add custom fasteners (not in CAD)
    add_fasteners()
    # Write all BOM files
    write_bom_files()
    # List all CAD objects by main and accent colors
    stl_files = get_stl_files()
    filename_results = get_filename_color_results(stl_files, cad_parts)
    write_filename_reports(filename_results)