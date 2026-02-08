#!/usr/bin/env python3
"""
Clean worksheet generator - reads Excel, generates HTML worksheets via Claude API
"""

import anthropic
import pandas as pd
from pathlib import Path
from docx import Document
from docx.shared import RGBColor
import re


class WorksheetGenerator:
    def __init__(self, api_key: str):
        self.client = anthropic.Anthropic(api_key=api_key)
        
    def read_lessons(self, excel_path: str) -> list[dict]:
        """Extract lessons from Excel file"""
        df = pd.read_excel(excel_path, header=4)
        lessons = []
        
        for _, row in df.iterrows():
            code = str(row.iloc[0]).strip()
            if code.startswith('C') and '.' in code:
                lessons.append({
                    'code': code,
                    'title': str(row.iloc[1]).strip(),
                    'knowledge': str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ''
                })
        
        return lessons
    
    def generate_worksheet(self, lesson: dict) -> str:
        """Generate HTML worksheet using Claude API"""
        
        prompt = f"""Create a GCSE Chemistry worksheet HTML file for Year 10 students.

LESSON: {lesson['title']}

CONTENT TO COVER:
{lesson['knowledge']}

REQUIREMENTS:
1. Create a complete HTML worksheet in the EXACT style I showed you before (with exit tickets, sections A/B/C)
2. Exit ticket: 3 multiple-choice questions with checkboxes
3. Section A - Knowledge Recall: Fill-in-blanks with word bank, matching activity, labelling diagram
4. Section B - Understanding: 8-10 questions testing understanding
5. Section C - Extended Response: 2-3 extended answer questions (4-6 marks each)
6. Use the EXACT CSS styling from the graphene worksheet
7. Include a visual diagram/image placeholder where relevant
8. Make questions engaging but strictly within AQA GCSE Chemistry specification
9. Base ALL questions on the lesson knowledge provided above

Return ONLY the complete HTML code, nothing else."""

        message = self.client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=16000,
            messages=[{"role": "user", "content": prompt}]
        )
        
        return message.content[0].text
    
    def save_worksheet(self, html: str, lesson: dict, output_dir: str = "worksheets"):
        """Save HTML worksheet to file"""
        Path(output_dir).mkdir(exist_ok=True)
        filename = f"{lesson['code']}_{lesson['title'].replace(' ', '_')}.html"
        filepath = Path(output_dir) / filename
        
        filepath.write_text(html, encoding='utf-8')
        print(f"✅ Created: {filename}")
        return filepath
    
    def save_word_doc(self, html: str, lesson: dict, output_dir: str = "worksheets"):
        """Save worksheet as Word document"""
        doc = Document()
        text = re.sub('<[^<]+?>', '', html)  # Strip HTML tags
        doc.add_paragraph(text)
        filename = f"{lesson['code']}_{lesson['title'].replace(' ', '_')}.docx"
        doc.save(Path(output_dir) / filename)
        print(f"✅ Created: {filename}")


def main():
    # Configuration
    API_KEY = "your-anthropic-api-key-here"
    EXCEL_FILE = "C4.1 Unit Overview.xlsx"
    
    # Initialize generator
    generator = WorksheetGenerator(API_KEY)
    
    # Read lessons
    print("📖 Reading lessons from Excel...")
    lessons = generator.read_lessons(EXCEL_FILE)
    print(f"Found {len(lessons)} lessons\n")
    
    # Generate worksheets for all lessons
    for lesson in lessons:
        print(f"🔄 Generating: {lesson['code']} - {lesson['title']}")
        
        try:
            html = generator.generate_worksheet(lesson)
            generator.save_worksheet(html, lesson)
            generator.save_word_doc(html, lesson)
        except Exception as e:
            print(f"❌ Error: {e}\n")
            continue
    
    print(f"\n✨ Done! Generated {len(lessons)} worksheets")


if __name__ == "__main__":
    main()