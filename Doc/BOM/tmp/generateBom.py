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
SNAKEOIL_PROJECT_PATH = str(BASE_PATH.parent.parent.parent)
CAD_FILE = Path(SNAKEOIL_PROJECT_PATH).joinpath('CAD/v1-180-assembly.FCStd')
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
    stl_files = STL.get_stl_files(SNAKEOIL_PROJECT_PATH, STL_PATH, EXCLUDE_DIRS, EXCLUDE_STRINGS)
    filename_results = CAD.get_filename_color_results(stl_files, cad_parts)
    STL.write_file_color_reports(filename_results)

if __name__ == '__main__':
    LOGGER.info(f"# Getting BOM from {CAD_FILE}")
    # Get assembly object from filepath
    try:
        cad_parts = CAD.get_cad_objects_from_cache()
    except KeyError:
        LOGGER.info("# No cached freecad_objects found. Reading from CAD file")
        # Open GUI if running from console, otherwise we know we are running from a macro
        if hasattr(Gui, 'showMainWindow'):
            Gui.showMainWindow()
        else:
            print("Running as macro")
        cad_assembly = App.open(str(CAD_FILE))
        cad_parts = CAD.get_cad_objects_from_freecad(cad_assembly)
        CAD.write_cad_objects_to_cache(cad_parts)
    # Generate bom-*.json files
    generate_bom(cad_parts)
    # Generate color-results-*.txt files
    generate_filename_color_reports(cad_parts)