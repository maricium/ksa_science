#!/usr/bin/env python3
"""
Unified Education System - Consolidates all education tools into one system
"""

import os
import sys
from pathlib import Path
import argparse

# Add all module paths
sys.path.append(str(Path(__file__).parent / "project-education"))
sys.path.append(str(Path(__file__).parent / "worksheet" / "worksheet"))
sys.path.append(str(Path(__file__).parent / "worksheet" / "mme"))

def generate_powerpoints():
    """Generate PowerPoint presentations using project-education"""
    print("🎯 Generating PowerPoint presentations...")
    try:
        from project_education.main import main as ppt_main
        ppt_main()
    except ImportError as e:
        print(f"❌ Error importing PowerPoint generator: {e}")

def generate_worksheets():
    """Generate worksheets using worksheet system"""
    print("📝 Generating worksheets...")
    try:
        from worksheet.main import main as worksheet_main
        worksheet_main()
    except ImportError as e:
        print(f"❌ Error importing worksheet generator: {e}")

def process_mme():
    """Process MME worksheets"""
    print("🧪 Processing MME worksheets...")
    try:
        from mme.main import main as mme_main
        mme_main()
    except ImportError as e:
        print(f"❌ Error importing MME processor: {e}")

def organize_all():
    """Organize all worksheets and resources"""
    print("📁 Organizing all resources...")
    try:
        from worksheet.organize_all_worksheets import main as organize_main
        organize_main()
    except ImportError as e:
        print(f"❌ Error importing organizer: {e}")

def clean_system():
    """Clean up redundant files and folders"""
    print("🧹 Cleaning up redundant files...")
    
    # Remove duplicate organization scripts
    duplicate_scripts = [
        "MSA/organize_mme_worksheets.py",
        "worksheet/worksheet/organize_worksheets.py",
        "worksheet/mme/organize_mme_worksheets.py"
    ]
    
    for script in duplicate_scripts:
        script_path = Path(script)
        if script_path.exists():
            script_path.unlink()
            print(f"🗑️  Removed: {script}")
    
    # Remove duplicate virtual environments
    venv_paths = [
        "foundation/venv",
        "worksheet/worksheet/venv"
    ]
    
    for venv in venv_paths:
        venv_path = Path(venv)
        if venv_path.exists():
            import shutil
            shutil.rmtree(venv_path)
            print(f"🗑️  Removed: {venv}")
    
    print("✅ Cleanup complete!")

def main():
    """Main unified system"""
    parser = argparse.ArgumentParser(description="Unified Education System")
    parser.add_argument("command", choices=[
        "powerpoints", "worksheets", "mme", "organize", "clean", "all"
    ], help="Command to run")
    
    args = parser.parse_args()
    
    print("🎓 Unified Education System")
    print("=" * 50)
    
    if args.command == "powerpoints":
        generate_powerpoints()
    elif args.command == "worksheets":
        generate_worksheets()
    elif args.command == "mme":
        process_mme()
    elif args.command == "organize":
        organize_all()
    elif args.command == "clean":
        clean_system()
    elif args.command == "all":
        print("🚀 Running all systems...")
        clean_system()
        generate_powerpoints()
        generate_worksheets()
        process_mme()
        organize_all()
        print("\n🎉 All systems completed!")
    
    print("\n✨ Done!")

if __name__ == "__main__":
    main()

