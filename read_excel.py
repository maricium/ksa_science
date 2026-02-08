"""
Compatibility shim for read_excel module.
All Excel reading functionality has been moved to readnow.excel_helpers.
This module maintains backward compatibility by re-exporting functions.
"""

import logging
from pathlib import Path

import pandas as pd

logger = logging.getLogger("read_excel")

# Import from excel_helpers for backward compatibility
try:
    import sys

    # Add project root to path if not already there
    project_root = Path(__file__).parent
    if str(project_root) not in sys.path:
        sys.path.insert(0, str(project_root))

    from readers.excel_reader import get_lesson_data

    # Re-export for backward compatibility
    __all__ = ["get_lesson_data", "view_all_excel_files"]
except ImportError:
    # Fallback if readnow.excel_helpers is not available
    logger.warning("Could not import from readnow.excel_helpers, using legacy implementation")

    def get_lesson_data(lesson_code, unit=None):
        """
        Extract lesson data (KNOW and DO) for a specific lesson code

        Args:
            lesson_code: e.g. "C4.1.4" or "C3.1.3"
            unit: e.g. "C4.1" (defaults to extracting from lesson_code)

        Returns:
            dict with 'code', 'title', 'know', 'do' or None if not found
        """
        # Determine which Excel file to use based on unit
        if not unit:
            unit = ".".join(lesson_code.split(".")[:2])  # Extract C4.1 from C4.1.4 or C3.1 from C3.1.3

        # Try to find Excel files in Lesson Resources/Unit Guidance
        lesson_resources = Path("Lesson Resources")
        if not lesson_resources.exists():
            lesson_resources = Path("../Lesson Resources")
        if not lesson_resources.exists():
            lesson_resources = Path("../../Lesson Resources")

        # Search for Excel files in Unit Guidance folders
        excel_files = list(lesson_resources.rglob(f"**/Unit Guidance/**/{unit}*.xlsx"))
        if not excel_files:
            excel_files = list(lesson_resources.rglob(f"**/Unit Guidance/**/*{unit}*.xlsx"))

        # Also try the old Excel_Files location as fallback
        if not excel_files:
            excel_file = Path(f"Excel_Files/{unit}_Unit_Overview.xlsx")
            if excel_file.exists():
                excel_files = [excel_file]
            else:
                excel_file = Path(f"../Excel_Files/{unit}_Unit_Overview.xlsx")
                if excel_file.exists():
                    excel_files = [excel_file]

        # Also check for C3.1 format in old location
        if not excel_files and unit.startswith("C3"):
            excel_file = Path("Excel_Files/C3.1 Unit Overview and Planning Proforma.xlsx")
            if excel_file.exists():
                excel_files = [excel_file]
            else:
                excel_file = Path("../Excel_Files/C3.1 Unit Overview and Planning Proforma.xlsx")
                if excel_file.exists():
                    excel_files = [excel_file]

        excel_file = excel_files[0] if excel_files else None  # Use first match

        if not excel_file or not excel_file.exists():
            logger.error(f"Excel file not found for unit {unit}")
            logger.error(f"Searched in: Lesson Resources/**/Unit Guidance/**/{unit}*.xlsx")
            return None

        logger.info(f"Reading Excel file: {excel_file.name}")
        logger.info(f"Unit: {unit}, Lesson code: {lesson_code}")

        # Try to determine sheet name for C4.2 files
        sheet_name = None
        if unit.startswith("C4.2"):
            try:
                xl_file = pd.ExcelFile(excel_file)
                # Look for sheet with "C4.2" in the name
                for sheet in xl_file.sheet_names:
                    if "C4.2" in sheet:
                        sheet_name = sheet
                        break
                if not sheet_name and xl_file.sheet_names:
                    sheet_name = xl_file.sheet_names[0]
            except:
                pass

        # C3.1 files have different structure - use header=3, lesson codes in column 1
        if unit.startswith("C3"):
            try:
                df = pd.read_excel(excel_file, sheet_name="C3.1 The Periodic Table", header=3)
                df = df.dropna(how="all")

                # Column 1 has lesson codes, Column 2 has title, Column 3 has "know" content
                for _idx, row in df.iterrows():
                    if len(row) > 1 and pd.notna(row.iloc[1]):
                        lesson_code_in_file = str(row.iloc[1]).strip()
                        if lesson_code_in_file == lesson_code:
                            # Column 2 has title, Column 3 has "know" content (bullet points)
                            title = str(row.iloc[2]).strip() if len(row) > 2 and pd.notna(row.iloc[2]) else lesson_code
                            know_content = str(row.iloc[3]).strip() if len(row) > 3 and pd.notna(row.iloc[3]) else ""
                            return {
                                "code": lesson_code,
                                "title": title,
                                "know": know_content,
                                "do": str(row.iloc[4]).strip() if len(row) > 4 and pd.notna(row.iloc[4]) else "",
                            }
            except Exception as e:
                logger.warning(f"Error reading C3.1 file: {e}")
                import traceback

                traceback.print_exc()

        # Also check if lesson_code itself starts with C3 (in case unit wasn't set correctly)
        if lesson_code.startswith("C3") and not unit.startswith("C3"):
            try:
                df = pd.read_excel(excel_file, sheet_name="C3.1 The Periodic Table", header=3)
                df = df.dropna(how="all")

                for _idx, row in df.iterrows():
                    if len(row) > 1 and pd.notna(row.iloc[1]):
                        lesson_code_in_file = str(row.iloc[1]).strip()
                        if lesson_code_in_file == lesson_code:
                            title = str(row.iloc[2]).strip() if len(row) > 2 and pd.notna(row.iloc[2]) else lesson_code
                            know_content = str(row.iloc[3]).strip() if len(row) > 3 and pd.notna(row.iloc[3]) else ""
                            return {
                                "code": lesson_code,
                                "title": title,
                                "know": know_content,
                                "do": str(row.iloc[4]).strip() if len(row) > 4 and pd.notna(row.iloc[4]) else "",
                            }
            except Exception:
                pass

        # Try reading with sheet name if we found one (for C4.2 files)
        if sheet_name:
            try:
                df = pd.read_excel(excel_file, sheet_name=sheet_name, skiprows=3)
                for _idx, row in df.iterrows():
                    # C4.2 format: lesson code is in column 1
                    if len(row) > 1 and pd.notna(row.iloc[1]):
                        code_col1 = str(row.iloc[1]).strip()
                        if code_col1 == lesson_code:
                            return {
                                "code": lesson_code,
                                "title": str(row.iloc[2]).strip() if len(row) > 2 and pd.notna(row.iloc[2]) else "",
                                "know": str(row.iloc[3]).strip() if len(row) > 3 and pd.notna(row.iloc[3]) else "",
                                "do": str(row.iloc[5]).strip() if len(row) > 5 and pd.notna(row.iloc[5]) else "",
                            }
            except Exception:
                pass

        # Try different skip rows to find the data (for C4 files and fallback)
        for skip in range(10):
            try:
                df = pd.read_excel(excel_file, skiprows=skip, sheet_name=sheet_name if sheet_name else None)
                first_col = str(df.columns[0]).lower().strip()

                if "lesson" in first_col and "code" in first_col:
                    # Check both column 0 and column 1 for lesson codes
                    for _idx, row in df.iterrows():
                        # Try column 1 first (C4.2 format - lesson code is in column 1)
                        if len(row) > 1 and pd.notna(row.iloc[1]):
                            code_col1 = str(row.iloc[1]).strip()
                            if code_col1 == lesson_code:
                                return {
                                    "code": lesson_code,
                                    "title": str(row.iloc[2]).strip() if len(row) > 2 and pd.notna(row.iloc[2]) else "",
                                    "know": str(row.iloc[3]).strip() if len(row) > 3 and pd.notna(row.iloc[3]) else "",
                                    "do": str(row.iloc[5]).strip() if len(row) > 5 and pd.notna(row.iloc[5]) else "",
                                }
                        # Try column 0 (other formats)
                        if len(row) > 0 and pd.notna(row.iloc[0]):
                            code_col0 = str(row.iloc[0]).strip()
                            if code_col0 == lesson_code:
                                return {
                                    "code": lesson_code,
                                    "title": str(row.iloc[1]).strip() if len(row) > 1 and pd.notna(row.iloc[1]) else "",
                                    "know": str(row.iloc[2]).strip() if len(row) > 2 and pd.notna(row.iloc[2]) else "",
                                    "do": str(row.iloc[4]).strip() if len(row) > 4 and pd.notna(row.iloc[4]) else "",
                                }
            except:
                continue

        return None


def view_all_excel_files():
    """Display summary of all Excel files in Lesson Resources/Unit Guidance"""
    # Search in new location first
    lesson_resources = Path("Lesson Resources")
    if not lesson_resources.exists():
        lesson_resources = Path("../Lesson Resources")
    if not lesson_resources.exists():
        lesson_resources = Path("../../Lesson Resources")

    excel_files = list(lesson_resources.rglob("**/Unit Guidance/**/*.xlsx"))

    # Fallback to old Excel_Files location
    if not excel_files:
        excel_files = list(Path("Excel_Files").rglob("*.xlsx"))
        if not excel_files:
            excel_files = list(Path("../Excel_Files").rglob("*.xlsx"))

    logger.info(f"Found {len(excel_files)} Excel files")

    for excel_file in excel_files:
        logger.info(f"{excel_file.name}")
        df = None

        for skip in range(10):
            try:
                temp_df = pd.read_excel(excel_file, skiprows=skip)
                first_col = str(temp_df.columns[0]).lower().strip()

                if "lesson" in first_col and "code" in first_col and str(temp_df.iloc[0, 0]).startswith("C"):
                    df = temp_df
                    break
            except:
                continue

        if df is not None:
            logger.info(f"{len(df)} lessons × {len(df.columns)} columns | {df.iloc[0, 0]} to {df.iloc[-1, 0]}")


if __name__ == "__main__":
    view_all_excel_files()
