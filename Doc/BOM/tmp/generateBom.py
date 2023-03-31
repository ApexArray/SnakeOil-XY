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
import sys
import FreeCAD as App  # type: ignore
import FreeCADGui as Gui  # type: ignore
import logging
from lib import BOM, STL, CAD

logging.basicConfig(
    # filename='generateBom.log', filemode='w', 
    format='%(levelname)s: %(message)s', level=logging.INFO
    )
LOGGER = logging.getLogger()

BASE_PATH = Path(os.path.dirname(__file__))
SNAKEOIL_PROJECT_PATH = BASE_PATH.parent.parent.parent
CAD_FILE = SNAKEOIL_PROJECT_PATH / 'CAD/v1-180-assembly.FCStd'
# Use EXTRA_CAD_FILES to find colors of parts not in the main assembly
EXTRA_CAD_FILES = [
    SNAKEOIL_PROJECT_PATH / 'WIP/component-assembly/toolhead-carrier-sherpa-1515-assembly.FCStd',
    SNAKEOIL_PROJECT_PATH / 'WIP/component-assembly/top-lid-assembly.FCStd',
    SNAKEOIL_PROJECT_PATH / 'WIP/component-assembly/sherpa-mini-assembly.FCStd',
    SNAKEOIL_PROJECT_PATH / 'WIP/component-assembly/bottom-panel-250-assembly.FCStd',
    SNAKEOIL_PROJECT_PATH / 'WIP/E-axis/sherpa-mini.FCStd',
]
STL_PATH = (
    SNAKEOIL_PROJECT_PATH / 'BETA3_Standard_Release_STL' / 'STLs').relative_to(SNAKEOIL_PROJECT_PATH)
# Ignore STL files in these directories
STL_EXCLUDE_DIRS = [
    (SNAKEOIL_PROJECT_PATH / "BETA3_Standard_Release_STL/STLs/Add-on"),
    (SNAKEOIL_PROJECT_PATH / "BETA3_Standard_Release_STL/STLs/Tools"),
    (SNAKEOIL_PROJECT_PATH / "BETA3_Standard_Release_STL/STLs/Panels/Bottom-panel/alt"),
    (SNAKEOIL_PROJECT_PATH / "BETA3_Standard_Release_STL/STLs/Z-axis/alt"),
]
# Convert to relative paths
STL_EXCLUDE_DIRS = [x.relative_to(SNAKEOIL_PROJECT_PATH) for x in STL_EXCLUDE_DIRS]
STL_EXCLUDE_STRINGS = [
    "OPTIONAL"
]

def generate_bom(cad_parts):
    """Builds bom from CAD objects and writes bom-*.json files"""
    BOM.get_bom_from_freecad_document(cad_parts)
    # Add custom fasteners (not in CAD)
    BOM.add_fasteners()
    # Write all BOM files
    BOM.write_bom_files()

def generate_filename_color_reports(cad_parts):
    """Builds filename color reports and writes to color-results-*.txt"""
    # List all CAD objects by main and accent colors
    stl_files = STL.get_stl_files(SNAKEOIL_PROJECT_PATH, STL_PATH, STL_EXCLUDE_DIRS, STL_EXCLUDE_STRINGS)
    filename_results = CAD.get_filename_color_results(stl_files, cad_parts)
    STL.write_file_color_reports(filename_results)
    return filename_results

if __name__ == '__main__':
    # Get assembly object from filepath
    cad_parts = CAD.get_cad_parts_from_file(CAD_FILE)
    extra_cad_parts = []
    for extra_cad_file in EXTRA_CAD_FILES:
        this_cad_parts = CAD.get_cad_parts_from_file(extra_cad_file)
        extra_cad_parts += this_cad_parts
    # Generate bom-*.json files
    generate_bom(cad_parts)
    # Generate color-results-*.txt files
    filename_results = generate_filename_color_reports(cad_parts + extra_cad_parts)