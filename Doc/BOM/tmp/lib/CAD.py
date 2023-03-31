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
from typing import Dict, List, Union
import FreeCAD as App  # type: ignore
import FreeCADGui as Gui  # type: ignore
import logging
from dataclasses import InitVar, dataclass, field
import shelve
import re
from difflib import SequenceMatcher

# Quick references to BOM part types.  Also provides type hinting in the BomPart dataclass
PRINTED_MAIN = "main"
PRINTED_ACCENT = "accent"
PRINTED_MISSING = "missing"
PRINTED_UNKNOWN_COLOR = "unknown"
PRINTED_CONFLICTING_COLORS = "conflicting"
FASTENER = "fastener"
OTHER = "other"

# Allowed types in a BOM
BOM_ITEM_TYPES = [
    PRINTED_MAIN, PRINTED_ACCENT, FASTENER, OTHER
]

BASE_PATH = Path(os.path.dirname(__file__)).parent

logging.basicConfig(
    # filename='generateBom.log', filemode='a', 
    format='%(levelname)s: %(message)s', level=logging.INFO
    )
LOGGER = logging.getLogger()

fastener_pattern = re.compile('.*-(Screw|Washer|HeatSet|Nut)')
revision_pattern = r'-.\d+$'  # Used to identify and strip revision numbers from CAD part names
# name cleaning patterns
revision_pattern = re.compile(r'-.\d+$')  # Used to identify and strip revision numbers from CAD part names
part_count_pattern = re.compile(r'^\d{1,2}x[_-]')  # Remove part counts at beginning of STL file names

def get_printed_part_color(part: App.DocumentObject):
    """Checks if CAD object is a printed part, as determined by its color (teal=main, blue=accent)

    Args:
        part (App.Part): FreeCAD object to check

    Returns:
        str: friendly name of color (ex: main, accent)
        None: if not a known color for printed parts
    """
    # Teal
    if part.ViewObject.ShapeColor == (0.3333333432674408, 1.0, 1.0, 0.0):  # type: ignore
        return PRINTED_MAIN
    # Blue
    elif part.ViewObject.ShapeColor == (0.6666666865348816, 0.6666666865348816, 1.0, 0.0):  # type: ignore
        return PRINTED_ACCENT
    else:
        return None

@dataclass
class BomItem:
    part: InitVar[App.DocumentObject]
    bom_item_type: str = field(init=False)  # What type of BOM entry (printed, fastener, other)
    name: str = field(init=False)  # We'll get the name from the part.label in the __post_init__ function
    clean_name: str = field(init=False)
    parent: str = field(init=False, default='')
    document: str = field(init=False, default='')
    color_category: str = field(init=False)
    raw_color: tuple = field(init=False)

    def __post_init__(self, part: App.DocumentObject):
        self.name = part.Label
        self.raw_color = part.ViewObject.ShapeColor  # type: ignore
        self.document = part.Document.Label
        self.clean_name = clean_name(self.name)
        # Remove numbers at end if they exist (e.g. 'M3-Washer004' becomes 'M3-Washer')
        while self.name[-1].isnumeric():
            self.name = self.name[:-1]
        if fastener_pattern.match(part.Label):
            self.bom_item_type = FASTENER
        else:
            color_category = get_printed_part_color(part)
            if color_category is not None:
                self.bom_item_type = color_category
            else:
                self.bom_item_type = OTHER
        # Add descriptive fastener names
        if self.bom_item_type == FASTENER:
            if hasattr(part, 'type'):
                if part.type == "ISO4762":  # type: ignore
                    self.name = f"Socket head {self.name}"
                if part.type == "ISO7380-1":  # type: ignore
                    self.name = f"Button head {self.name}"
                if part.type == "ISO4026": # type: ignore
                    self.name = f"Grub {self.name}"
                if part.type == "ISO4032":  # type: ignore
                    self.name = f"Hex {self.name}"
                if part.type == "ISO7092":  # type: ignore
                    self.name = f"Small size {self.name}"
                if part.type == "ISO7093-1":  # type: ignore
                    self.name = f"Big size {self.name}"
                if part.type == "ISO7089":  # type: ignore
                    self.name = f"Standard size {self.name}"
                if part.type == "ISO7090":  # type: ignore
                    self.name = f"Standard size {self.name}"
        # Try to add parent object
        try:
            self.parent = part.Parents[0][0].Label
            if self.parent == "snakeoilxy-180":
                self.parent = part.Parents[1][0].Label
        except:
            pass
    
    def __str__(self) -> str:
        return f"{self.document}: {self.parent}/{self.name}"
    
    def __repr__(self) -> str:
        return self.__str__()
    
def _get_cad_parts_from_cache(var_name: str) -> List[BomItem]:
    with shelve.open('cad_cache') as db:
        return db[var_name]
    
def _write_cad_parts_to_cache(var_name: str, cad_parts: List[BomItem]):
    LOGGER.info(f"Writing {var_name} CAD parts to cache")
    with shelve.open('cad_cache') as db:
        db[var_name] = cad_parts

def _get_cad_parts_from_freecad_assembly(assembly: App.Document) -> List[BomItem]:
    freecad_objects = None
    # Try to read from cache to avoid length process of reading CAD file
    LOGGER.debug("# Getting parts from", assembly.Label)
    freecad_objects = [BomItem(x) for x in assembly.Objects if x.TypeId.startswith('Part::')]
    # Recurse through each linked file
    for linked_file in assembly.findObjects("App::Link"):
        LOGGER.debug("# Getting linked parts from", linked_file.Name)
        freecad_objects += _get_cad_parts_from_freecad_assembly(
            linked_file.LinkedObject.Document)  # type: ignore
    return freecad_objects

def get_cad_parts_from_file(path: Path, use_cache=True) -> List[BomItem]:
    if use_cache:
        try:
            return _get_cad_parts_from_cache(path.name)
        except KeyError:
            LOGGER.debug(f"No cache for {path}")
    LOGGER.info(f"Loading CAD parts from {path}")
    Gui.showMainWindow()
    cad_assembly = App.open(str(path))
    cad_parts = _get_cad_parts_from_freecad_assembly(cad_assembly)
    _write_cad_parts_to_cache(path.name, cad_parts)
    return cad_parts

def clean_name(name: str):
    name = name.replace('_', '-')
    name = name.replace('.stl', '')
    for pattern in [revision_pattern, part_count_pattern]:
        matches = re.findall(pattern, name)
        if matches:
            if len(matches) > 1:
                print("ERROR: multiple regex matches found")
                print(f"  {'  '.join(matches)}")
            name = name.replace(matches[0], '')
    # Remove numbers at end if they exist (e.g. 'M3-Washer004' becomes 'M3-Washer')
    while name[-1].isnumeric():
        name = name[:-1]
    return name

def search_cad_objects__cad_part_name_in_filename(file_name: str, cad_objects: List[BomItem]):
    """Returns list of CAD objects if cad part name in file_name"""
    file_name = clean_name(file_name)
    return [part for part in cad_objects if part.clean_name in file_name]

def search_cad_objects__filename_in_cad_part_name(file_name: str, cad_objects: List[BomItem]):
    """Returns list of CAD objects if file_name is in cad part name"""
    file_name = clean_name(file_name)
    return [part for part in cad_objects if file_name in part.clean_name]

def search_cad_objects__fuzzy(file_name: str, cad_objects: List[BomItem], cutoff_ratio=0.9):
    file_name = clean_name(file_name)
    return [part for part in cad_objects if (
        SequenceMatcher(None, file_name, part.clean_name).ratio() >= cutoff_ratio
        )]

def search_cad_objects__fuzzy_top_result(file_name: str, cad_objects: List[BomItem]):
    file_name = clean_name(file_name)
    top_results = []
    top_ratio = 0.0
    for part in cad_objects:
        ratio = SequenceMatcher(None, file_name, part.clean_name).ratio()
        # Reset list with this part as the sole member if this is the best ratio we've seen
        if ratio > top_ratio:
            top_results = [part]
            top_ratio = ratio
        # Append part to list if tied with top_ratio
        elif ratio == top_ratio:
            top_results.append(part)
            top_ratio = ratio
    return top_results

def get_part_color_from_filename(file_name: str, cad_objects: List[BomItem]) -> str:
    """Check if part_name is a main or accent color. Returns 'main', 'accent', 0 or obj containing error info"""
    # Find objects in each list with names container in our filename
    all_results = search_cad_objects__cad_part_name_in_filename(file_name, cad_objects)
    if not all_results:
        LOGGER.debug(f"Trying alternative search method for {file_name}")
        all_results = search_cad_objects__filename_in_cad_part_name(file_name, cad_objects)
        if all_results:
            LOGGER.debug(f"Found {len(all_results)} matches using alternative serach method for {file_name}")
        else:
            LOGGER.debug(f"Trying fuzzy search method for {file_name}")
            all_results = search_cad_objects__fuzzy(file_name, cad_objects, 0.8)
            if all_results:
                LOGGER.warning(f"Found fuzzy matches {file_name} -> {all_results}")
            else:
                fuzzy_results = search_cad_objects__fuzzy_top_result(file_name, cad_objects)
                if fuzzy_results:
                    LOGGER.warning(f"{file_name} top fuzzy matches {fuzzy_results}")
    main_results = [part for part in all_results if part.bom_item_type == PRINTED_MAIN]
    accent_results = [part for part in all_results if part.bom_item_type == PRINTED_ACCENT]
    unknown_results = [part for part in all_results if part.bom_item_type not in [PRINTED_MAIN, PRINTED_ACCENT]]
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
            LOGGER.debug(full_report + main_color_report)
            return PRINTED_MAIN
        elif total_colored_count == accent_count:
            LOGGER.debug(full_report + accent_color_report)
            return PRINTED_ACCENT
        else:
            raise Exception("Total color count does not match main_count or accent_count")
    # Display error if we found matching results with both main and accent colors
    else:
        msg = f"{PRINTED_CONFLICTING_COLORS} colors found:\n" + main_color_report + accent_color_report
        LOGGER.error(f"{file_name} {msg}")
        return msg

def get_filename_color_results(stl_files: List[Path], cad_parts: List[BomItem]) -> Dict[str, Union[List[Path], List[str]]]:
    """return dictionary of filename['main'|'accent'|0|obj]"""
    # main_colors, accent_colors = load_printed_parts_from_file()
    file_results = {
        file_path: get_part_color_from_filename(file_path.name, cad_parts) for file_path in stl_files
        }
    main_parts = [fp for (fp, result) in file_results.items() if result == PRINTED_MAIN]
    accent_parts = [fp for (fp, result) in file_results.items() if result == PRINTED_ACCENT]
    missing_parts = [fp for (fp, result) in file_results.items() if result.startswith(PRINTED_MISSING)]
    unknown_color_parts = [
        f"{fp.as_posix()} {result}" for (fp, result) in file_results.items() if result.startswith(PRINTED_UNKNOWN_COLOR)
        ]
    conflicting_parts = [
        f"{fp.as_posix()} {result}" for (fp, result) in file_results.items() if result.startswith(PRINTED_CONFLICTING_COLORS)
        ]
    msg = f"""# Total STL files: {len(stl_files)}
# Total main parts: {len(main_parts)}
# Total accent parts: {len(accent_parts)}
# Total missing parts: {len(missing_parts)}
# Total unknown colored parts: {len(unknown_color_parts)}
# Total conflicting parts: {len(conflicting_parts)}"""
    LOGGER.info("\n"+msg)
    with open(BASE_PATH / 'color-results-overview.txt', 'w') as file:
        file.write(msg)
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