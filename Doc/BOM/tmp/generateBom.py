"""
Run with FreeCAD's bundled interpreter, or as a FreeCAD macro
"""
import os
FREECADPATH = os.getenv('FREECADPATH', '/usr/lib/freecad-python3/lib/')
FREECADPATH = os.getenv('FREECADPATH', 'C:/Program Files/FreeCAD 0.20/bin')
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
from dataclasses import dataclass, field
from enum import Enum
import shelve
from lib.CAD import (
    BomItem, PRINTED_MAIN, PRINTED_ACCENT, PRINTED_MISSING, PRINTED_UNKNOWN_COLOR, PRINTED_CONFLICTING_COLORS, 
    FASTENER, OTHER, get_cad_objects_from_freecad, get_cad_objects_from_cache, write_cad_objects_to_cache
    )

logging.basicConfig(
    filename='generateBom.log', filemode='a', format='%(levelname)s: %(message)s', level=logging.INFO
    )
LOGGER = logging.getLogger()

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

def get_part_color_from_filename(file_name: str, cad_objects: List[BomItem]):
    """Check if part_name is a main or accent color. Returns 'main', 'accent', 0 or obj containing error info"""
    # Find objects in each list with names container in our filename
    all_results = [part for part in cad_objects if part.name in file_name]
    main_results = [part for part in all_results if part.type == PRINTED_MAIN]
    accent_results = [part for part in all_results if part.type == PRINTED_ACCENT]
    unknown_results = [part for part in all_results if part.type not in [PRINTED_MAIN, PRINTED_ACCENT]]
    # Help variables for logging below
    main_count = len(main_results)
    accent_count = len(accent_results)
    total_colored_count = main_count + accent_count
    main_list = '\n'.join([f'    - {part}' for part in main_results])
    accent_list = '\n'.join([f'    - {part}' for part in accent_results])
    main_color_report = f"  main colors:\n{main_list}\n"
    accent_color_report = f"  accent colors:\n{accent_list}\n"
    # Ideally, we should have exactly one match.
    if total_colored_count == 1:
        if main_count == 1:
            return PRINTED_MAIN
        elif accent_count == 1:
            return PRINTED_ACCENT
        else:
            raise ValueError(f"total count is {total_colored_count}, main and accent count != 1")
    # Return 0 if no results were found
    elif total_colored_count == 0:
        if len(unknown_results) > 0:
            msg = f"{PRINTED_UNKNOWN_COLOR} colors found:\n"
            msg += str(unknown_results)
            LOGGER.error(f"{file_name} {msg}")
            return msg
        else:
            return PRINTED_MISSING
    # Proceed with warning if more than one match, but all matches are either main OR accent color
    elif total_colored_count > 1 and total_colored_count in [main_count, accent_count]:
        full_report = f"{file_name} matches multiple CAD objects of the same color:\n"
        if total_colored_count == main_count:
            LOGGER.warning(full_report + main_color_report)
            return PRINTED_MAIN
        elif total_colored_count == accent_count:
            LOGGER.warning(full_report + accent_color_report)
            return PRINTED_ACCENT
    # Display error if we found matching results with both main and accent colors
    else:
        msg = f"{PRINTED_CONFLICTING_COLORS} colors found:\n" + main_color_report + accent_color_report
        LOGGER.error(f"{file_name} {msg}")
        return msg

def get_filename_color_results(printed_parts: List[BomItem]):
    """return dictionary of filename['main'|'accent'|0|obj]"""
    stl_files = get_stl_files()
    # main_colors, accent_colors = load_printed_parts_from_file()
    file_results = {
        file_path: get_part_color_from_filename(file_path.name, printed_parts) for file_path in stl_files
        }
    main_parts = [fp for (fp, result) in file_results.items() if result == PRINTED_MAIN]
    accent_parts = [fp for (fp, result) in file_results.items() if result == PRINTED_ACCENT]
    missing_parts = [fp for (fp, result) in file_results.items() if result.startswith(PRINTED_MISSING)]
    unknown_color_parts = [
        f"{fp} {result}" for (fp, result) in file_results.items() if result.startswith(PRINTED_UNKNOWN_COLOR)
        ]
    conflicting_parts = [
        f"{fp} {result}" for (fp, result) in file_results.items() if result.startswith(PRINTED_CONFLICTING_COLORS)
        ]
    LOGGER.info(f"# Total STL files: {len(stl_files)}")
    LOGGER.info(f"# Total main parts: {len(main_parts)}")
    LOGGER.info(f"# Total accent parts: {len(accent_parts)}")
    LOGGER.info(f"# Total missing parts: {len(missing_parts)}")
    LOGGER.info(f"# Total unknown colored parts: {len(unknown_color_parts)}")
    LOGGER.info(f"# Total conflicting parts: {len(conflicting_parts)}")
    assert len(stl_files) == sum([
        len(main_parts), len(accent_parts), len(missing_parts), 
        len(unknown_color_parts), len(conflicting_parts)
    ])
    return {
        PRINTED_MAIN: main_parts, 
        PRINTED_ACCENT: accent_parts, 
        PRINTED_MISSING: missing_parts,
        PRINTED_UNKNOWN_COLOR: unknown_color_parts,
        PRINTED_CONFLICTING_COLORS: conflicting_parts
        }

def write_filename_reports(filename_results):
    for category in [
        PRINTED_MAIN, PRINTED_ACCENT, PRINTED_MISSING, PRINTED_UNKNOWN_COLOR, PRINTED_CONFLICTING_COLORS
        ]:
        with open(f'{category}-color.log', 'w') as file:
            results: list[Path] = filename_results[category]
            formatted_results: list[str] = [str(x) for x in results]
            sorted_results = sorted(formatted_results)
            file.write('\n'.join(sorted_results))

if __name__ == '__main__':
    LOGGER.info(f"# Getting BOM from {target_file}")
    # Get assembly object from filepath
    try:
        cad_parts = get_cad_objects_from_cache()
    except KeyError:
        LOGGER.info("# No cached freecad_objects found. Reading from CAD file")
        cad_assembly = App.open(str(target_file))
        cad_parts = get_cad_objects_from_freecad(cad_assembly)
        write_cad_objects_to_cache(cad_parts)
    get_bom_from_freecad_document(cad_parts)
    # Add custom fasteners (not in CAD)
    add_fasteners()
    # Write all BOM files
    write_bom_files()
    # List all CAD objects by main and accent colors
    filename_results = get_filename_color_results(cad_parts)
    write_filename_reports(filename_results)