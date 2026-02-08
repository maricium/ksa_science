#!/usr/bin/env python3
"""
Consolidated script to organize all worksheets by unit/topic
"""

import os
import shutil
import re
from pathlib import Path

def organize_worksheets_by_unit(worksheets_dir, folder_name="worksheets"):
    """Organize worksheets by unit (C4.1, C4.2, C4.3, C4.4)"""
    
    if not worksheets_dir.exists():
        print(f"❌ {folder_name} directory not found")
        return False
    
    # Create unit folders
    units = ["C4.1", "C4.2", "C4.3", "C4.4"]
    for unit in units:
        unit_dir = worksheets_dir / unit
        unit_dir.mkdir(exist_ok=True)
        print(f"📁 Created folder: {unit}")
    
    # Move worksheets to appropriate unit folders
    moved_count = 0
    for file_path in worksheets_dir.iterdir():
        if file_path.is_file() and file_path.suffix == '.docx':
            filename = file_path.name
            
            # Extract unit code from filename (e.g., C4.1.1_... -> C4.1)
            match = re.match(r'(C4\.\d+)\.', filename)
            if match:
                unit_code = match.group(1)
                if unit_code in units:
                    # Move file to unit folder
                    destination = worksheets_dir / unit_code / filename
                    shutil.move(str(file_path), str(destination))
                    print(f"📄 Moved: {filename} -> {unit_code}/")
                    moved_count += 1
                else:
                    print(f"⚠️  Unknown unit: {unit_code} for {filename}")
            else:
                print(f"⚠️  Could not extract unit from: {filename}")
    
    print(f"✅ {folder_name} reorganization complete! Moved {moved_count} worksheets")
    return True

def organize_mme_worksheets_by_topic(worksheets_dir):
    """Organize MME worksheets by chemistry topics"""
    
    if not worksheets_dir.exists():
        print("❌ MME worksheets directory not found")
        return False
    
    # Define topic categories based on chemistry curriculum
    topics = {
        "Atomic_Structure": [
            "Structure_of_an_Atom", "The_Development_of_the_Model_of_the_Atom", 
            "The_Development_of_the_Periodic_Table", "The_Periodic_Table"
        ],
        "Bonding": [
            "Chemical_Bonds_Ionic", "Chemical_Bonds_Covalent", "Chemical_Bonds_Metallic",
            "Structure_and_Bonding_of_Carbon", "Diamond", "Graphite", "Graphene"
        ],
        "Acids_Bases": [
            "Acids_and_PH", "Reaction_of_acids", "Titrations"
        ],
        "Electrolysis": [
            "Electrolysis", "Chemical_Cells_and_Fuel_Cells"
        ],
        "Energy": [
            "Energy_Transfers", "Yield_and_Atom_Economy"
        ],
        "Organic_Chemistry": [
            "Crude_Oil_and_Cracking", "Alkenes_Alcohols_and_Carboxylic_Acids", 
            "Polymers", "Nanoparticles"
        ],
        "Rates_Equilibrium": [
            "Rate_of_Reactions", "Equilibrium", "Haber_Process"
        ],
        "Atmosphere": [
            "Earths_Atmosphere", "Greenhouse_Gases"
        ],
        "Materials": [
            "Cermaics_Polymers_and_Composites", "Materials_and_Resources", 
            "Reactivity_of_Metals"
        ],
        "Analysis": [
            "Tests_for_Ions_and_Gases", "Purity_and_Chromatography", 
            "Chemical_analysis"
        ],
        "Moles_Calculations": [
            "Moles", "Chemical_Reactions_and_Relative_Formula"
        ],
        "Water_Treatment": [
            "Water_Waste_and_Treatment"
        ],
        "Other": [
            "Modern_Slavery_Statement", "Safeguarding", "Worksheet_Answers", "Worksheet_Question"
        ]
    }
    
    # Create topic folders
    for topic in topics.keys():
        topic_dir = worksheets_dir / topic
        topic_dir.mkdir(exist_ok=True)
        print(f"📁 Created folder: {topic}")
    
    # Move worksheets to appropriate topic folders
    moved_count = 0
    for file_path in worksheets_dir.iterdir():
        if file_path.is_file() and file_path.suffix == '.docx':
            filename = file_path.name
            moved = False
            
            # Try to match filename to topics
            for topic, keywords in topics.items():
                for keyword in keywords:
                    if keyword.lower() in filename.lower():
                        # Move file to topic folder
                        destination = worksheets_dir / topic / filename
                        shutil.move(str(file_path), str(destination))
                        print(f"📄 Moved: {filename} -> {topic}/")
                        moved_count += 1
                        moved = True
                        break
                if moved:
                    break
            
            if not moved:
                # Move to Other folder if no match found
                destination = worksheets_dir / "Other" / filename
                shutil.move(str(file_path), str(destination))
                print(f"📄 Moved: {filename} -> Other/")
                moved_count += 1
    
    print(f"✅ MME worksheets reorganization complete! Moved {moved_count} worksheets")
    return True

def main():
    print("🔄 Organizing all worksheets...")
    
    # Organize regular worksheets by unit
    print("\n📚 Organizing regular worksheets by unit...")
    regular_worksheets = Path("worksheet/worksheets")
    organize_worksheets_by_unit(regular_worksheets, "Regular worksheets")
    
    # Organize MME worksheets by topic
    print("\n🧪 Organizing MME worksheets by topic...")
    mme_worksheets = Path("mme/word_documents/worksheets")
    organize_mme_worksheets_by_topic(mme_worksheets)
    
    print("\n🎉 All worksheets organized successfully!")

if __name__ == "__main__":
    main()

