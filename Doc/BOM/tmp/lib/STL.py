"""
Run with FreeCAD's bundled interpreter, or as a FreeCAD macro
"""
import logging
import os
if os.name == 'nt':
    FREECADPATH = os.getenv('FREECADPATH', 'C:/Program Files/FreeCAD 0.20/bin')
else:
    FREECADPATH = os.getenv('FREECADPATH', '/usr/lib/freecad-python3/lib/')
import sys
sys.path.append(FREECADPATH)
from pathlib import Path
import sys
from typing import Dict, List, Union
from lib import CAD

logging.basicConfig(
    # filename='generateBom.log', filemode='w', 
    format='%(levelname)s: %(message)s', level=logging.INFO
    )
LOGGER = logging.getLogger()

BASE_PATH = Path(os.path.dirname(__file__)).parent

def filter_stl_files(stl_files: List[Path], exclude_dirs: List[Path] = [], exclude_strings: List[str] = []):
    filtered_stl_files = []
    for file in stl_files:
        keep = True
        for excluded_fp in exclude_dirs:
            if str(file).startswith(str(excluded_fp)):
                keep = False
        for excluded_str in exclude_strings:
            if excluded_str.lower() in str(file).lower():
                keep = False
        if keep == True:
            filtered_stl_files.append(file)
    return filtered_stl_files

def get_stl_files(snakeoil_project_path: Path, target_stl_path: str, 
                  exclude_dirs: List[Path] = [], exclude_strings: List[str] = []) -> List[Path]:
    dir = Path(snakeoil_project_path)
    stl_files = [x.relative_to(dir) for x in dir.glob(f"{target_stl_path}/**/*.stl")]
    stl_files = filter_stl_files(stl_files, exclude_dirs, exclude_strings)
    LOGGER.info(f"# Found {len(stl_files)} stl files")
    return stl_files

def write_file_color_reports(filename_results: Dict[str, List[Union[Path, str]]]):
    for category in [
        CAD.PRINTED_MAIN, CAD.PRINTED_ACCENT, CAD.PRINTED_MISSING, 
        CAD.PRINTED_UNKNOWN_COLOR, CAD.PRINTED_CONFLICTING_COLORS
        ]:
        with open(BASE_PATH / f'color-results-{category}.txt', 'w') as file:
            results: list[Path] = filename_results[category]
            if issubclass(type(results[0]), Path):
                results: list[str] = [x.as_posix() for x in results]
            sorted_results = sorted(results)
            file.write('\n'.join(sorted_results))