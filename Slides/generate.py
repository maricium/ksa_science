"""Generate PowerPoint slides from Excel lesson data"""

import os
import re
from pathlib import Path
from lib.processor import create_slide_from_template
from lib.reader import read_lesson_data
from lib.utils import find_path
import config


def _get_excel_files():
    """Find all unit overview Excel files"""
    lesson_resources = find_path(config.LESSON_RESOURCES_DIR)
    excel_files = list(lesson_resources.rglob("**/Unit Guidance/**/*.xlsx"))
    # Filter out temporary Excel lock files (starting with ~$)
    excel_files = [f for f in excel_files if not Path(f).name.startswith('~$')]
    if not excel_files:
        print("❌ No unit overview Excel files found")
        return []
    print(f"Found {len(excel_files)} unit overview Excel files")
    return excel_files


def _get_sheet_name(unit_code, excel_file):
    """Auto-detect sheet name from unit code"""
    try:
        import pandas as pd
        excel = pd.ExcelFile(excel_file)
        unit_clean = unit_code.replace('.', '')
        for sheet in excel.sheet_names:
            if unit_clean in sheet.replace(' ', '').replace('.', '').replace('-', '') or sheet.replace(' ', '').replace('.', '').replace('-', '').startswith(unit_clean):
                return sheet
        return excel.sheet_names[0] if excel.sheet_names else config.DEFAULT_SHEET
    except:
        return config.DEFAULT_SHEET


def main():
    """Main entry point"""
    print("🚀 Slide Generator\n" + "=" * 60)
    os.makedirs(config.OUTPUT_DIR, exist_ok=True)
    print(f"📁 Output: {config.OUTPUT_DIR}\n")
    
    template_path = Path("templates") / config.TEMPLATE_NAME
    if not template_path.exists():
        template_path = find_path(f"templates/{config.TEMPLATE_NAME}")
    if not template_path.exists():
        print(f"❌ Template not found: {config.TEMPLATE_NAME}")
        return
    
    units_processed = []
    for excel_file in _get_excel_files():
        unit_code = re.search(r'([CPB]\d+\.\d+)', str(excel_file))
        if not unit_code or (config.TARGET_UNIT and unit_code.group(1) != config.TARGET_UNIT):
            continue
        unit_code = unit_code.group(1)
        
        print(f"\n{'='*60}\nProcessing Unit: {unit_code}\n{'='*60}")
        try:
            df = read_lesson_data(str(excel_file), _get_sheet_name(unit_code, str(excel_file)))
            print(f"Found {len(df)} lessons")
            if len(df) == 0:
                continue
            
            for _, row in df.iterrows():
                lesson_data = {
                    'lesson_code': str(row.get('lesson_code', '')).strip(),
                    'lesson_title': str(row.get('lesson_title', '')).strip(),
                    'knowledge_objectives': str(row.get('knowledge_objectives', '')).strip(),
                    'skill_objectives': str(row.get('skill_objectives', '')).strip(),
                    'exit_ticket': str(row.get('exit_ticket', '')).strip(),
                }
                if not lesson_data['lesson_code'] or lesson_data['lesson_code'] == 'nan' or (config.TARGET_LESSON and lesson_data['lesson_code'].upper() != config.TARGET_LESSON.upper()):
                    continue
                
                print(f"\nProcessing: {lesson_data['lesson_code']} - {lesson_data['lesson_title']}")
                clean_title = re.sub(r'[<>:"/\\|?*]', '', lesson_data['lesson_title']).strip()[:60]
                filename = f"{lesson_data['lesson_code']} {clean_title}.pptx" if clean_title else f"{lesson_data['lesson_code']}.pptx"
                output_path = Path(config.OUTPUT_DIR) / filename
                
                if create_slide_from_template(str(template_path), lesson_data, str(output_path)):
                    print(f"✅ Created: {lesson_data['lesson_code']}")
                else:
                    print(f"❌ Failed: {lesson_data['lesson_code']}")
            units_processed.append(unit_code)
        except Exception as e:
            print(f"❌ Error: {e}")
            import traceback
            traceback.print_exc()
    
    print(f"\n{'='*60}\n✅ Completed! Generated slides for: {units_processed}\n📁 Output: {config.OUTPUT_DIR}/")


if __name__ == "__main__":
    main()
