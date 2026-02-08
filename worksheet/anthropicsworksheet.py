#!/usr/bin/env python3
"""
Worksheet Generator - Reads Excel curriculum and generates custom HTML worksheets
Based on ACTUAL "What students need to KNOW" and "What students need to DO" columns
"""

import pandas as pd
from pathlib import Path
import re
from docx import Document


class WorksheetGenerator:
    """Generate custom worksheets based on lesson knowledge and skills"""
    
    def __init__(self, excel_path: str):
        """Load curriculum from Excel file"""
        # Read Excel, skip first few rows to get to the data
        self.df = pd.read_excel(excel_path, header=4)
        
        # Column mapping (adjust indices if needed based on your Excel structure)
        self.col_code = 0      # Lesson code
        self.col_title = 1     # Lesson title  
        self.col_know = 2      # What students need to KNOW
        self.col_do = 5        # What students need to DO
        
        print(f"✅ Loaded curriculum from {excel_path}")
        print(f"📊 Found {len(self.df)} rows")
    
    def get_lesson(self, lesson_code: str) -> dict:
        """Extract lesson data by code (e.g., 'C4.1.12')"""
        for idx, row in self.df.iterrows():
            if str(row.iloc[self.col_code]).strip() == lesson_code:
                return {
                    'code': lesson_code,
                    'title': str(row.iloc[self.col_title]).strip(),
                    'knowledge': str(row.iloc[self.col_know]) if pd.notna(row.iloc[self.col_know]) else '',
                    'skills': str(row.iloc[self.col_do]) if pd.notna(row.iloc[self.col_do]) else ''
                }
        
        print(f"❌ Lesson {lesson_code} not found")
        return None
    
    def analyze_skills(self, skills_text: str) -> dict:
        """Analyze the DO column to determine what activities to include"""
        skills_lower = skills_text.lower()
        
        analysis = {
            'needs_calculations': any(word in skills_lower for word in ['calculate', 'work out', 'mass', 'volume', 'percentage']),
            'needs_diagrams': any(word in skills_lower for word in ['draw', 'diagram', 'figure', 'show', 'dot and cross', 'structure']),
            'needs_graphs': any(word in skills_lower for word in ['plot', 'bar chart', 'graph', 'interpret']),
            'needs_definitions': any(word in skills_lower for word in ['what is meant by', 'define', 'state what', 'recall']),
            'needs_explanations': any(word in skills_lower for word in ['explain', 'why', 'describe']),
            'needs_comparisons': any(word in skills_lower for word in ['compare', 'difference', 'contrast']),
            'is_bonding_topic': any(word in skills_lower for word in ['covalent', 'ionic', 'dot and cross', 'electron']),
            'is_polymer_topic': any(word in skills_lower for word in ['polymer', 'repeating unit', 'thermosoftening']),
            'has_data_tasks': any(word in skills_lower for word in ['table', 'data', 'interpret', 'chart'])
        }
        
        return analysis
    
    def generate_section_a(self, lesson: dict, analysis: dict) -> str:
        """Generate Section A based on skills analysis"""
        knowledge_points = [k.strip() for k in lesson['knowledge'].split('\n') if k.strip()]
        skills_list = [s.strip() for s in lesson['skills'].split('\n') if s.strip()]
        
        # Extract key terms for word bank
        keywords = self.extract_keywords(lesson['knowledge'])
        word_bank = ', '.join(keywords[:12])
        
        section_a = ""
        
        # Part 1: Always start with definitions/recall if needed
        if analysis['needs_definitions']:
            section_a += self.generate_definitions_part(skills_list)
        else:
            section_a += self.generate_fill_blanks(knowledge_points, word_bank)
        
        # Part 2: Diagrams, data handling, or matching
        if analysis['needs_diagrams'] and analysis['is_bonding_topic']:
            section_a += self.generate_bonding_diagrams(lesson['title'])
        elif analysis['needs_diagrams'] and analysis['is_polymer_topic']:
            section_a += self.generate_polymer_diagrams()
        elif analysis['has_data_tasks'] or analysis['needs_graphs']:
            section_a += self.generate_data_handling_part()
        else:
            section_a += self.generate_matching_activity(knowledge_points)
        
        # Part 3: Calculations or additional practice
        if analysis['needs_calculations']:
            section_a += self.generate_calculations_part(skills_list)
        elif analysis['needs_comparisons']:
            section_a += self.generate_comparison_table(lesson['title'])
        else:
            section_a += self.generate_labelling_activity()
        
        return section_a
    
    def generate_definitions_part(self, skills_list: list) -> str:
        """Generate Part 1: Definitions and recall questions"""
        # Extract actual definition questions from skills
        definition_qs = [s for s in skills_list if 'what is' in s.lower() or 'define' in s.lower() or 'state what' in s.lower()]
        
        questions_html = ""
        for i, q in enumerate(definition_qs[:3], 1):
            questions_html += f"""
    <p><strong>{i}. {q}</strong></p>
    <div class="answer-space"></div>
"""
        
        return f"""
    <h4>Part 1 – Definitions and Key Facts</h4>
{questions_html}
"""
    
    def generate_fill_blanks(self, knowledge_points: list, word_bank: str) -> str:
        """Generate fill-in-the-blanks activity"""
        blanks_html = ""
        for i, point in enumerate(knowledge_points[:6], 1):
            # Create blank by replacing first significant word
            words = point.split()
            if len(words) > 3:
                blank_sentence = point.replace(words[2], "_____", 1)
            else:
                blank_sentence = point
            blanks_html += f"""        <li>{blank_sentence}</li>\n"""
        
        return f"""
    <h4>Part 1 – Fill in the Blanks</h4>
    
    <div class="word-bank">
        <strong>Word Bank:</strong><br>
        <em>{word_bank}</em>
    </div>
    
    <ol class="fill-blank">
{blanks_html}
    </ol>
"""
    
    def generate_bonding_diagrams(self, title: str) -> str:
        """Generate dot-and-cross diagram practice"""
        if 'covalent' in title.lower():
            molecules = ['H₂', 'Cl₂', 'H₂O', 'CH₄', 'O₂', 'CO₂']
        elif 'ionic' in title.lower():
            molecules = ['NaCl', 'MgO', 'CaCl₂']
        else:
            molecules = ['Example 1', 'Example 2', 'Example 3']
        
        boxes_html = ""
        for i, molecule in enumerate(molecules[:3], 1):
            boxes_html += f"""
    <div class="drawing-box">
        <h4>{i}. Draw the dot-and-cross diagram for {molecule}</h4>
        <div class="drawing-space"></div>
    </div>
"""
        
        return f"""
    <h4>Part 2 – Drawing Practice</h4>
    <p><strong>Draw dot-and-cross diagrams for the following. Show all outer shell electrons.</strong></p>
    
    <div class="drawing-instructions">
        <strong>Remember:</strong>
        <ul>
            <li>Use dots (•) for electrons from one atom</li>
            <li>Use crosses (×) for electrons from the other atom</li>
            <li>Show only outer shell electrons</li>
            <li>Draw circles to represent atoms</li>
        </ul>
    </div>
{boxes_html}
"""
    
    def generate_polymer_diagrams(self) -> str:
        """Generate polymer repeating unit practice"""
        return """
    <h4>Part 2 – Drawing Repeating Units</h4>
    <p><strong>Draw the repeating units for the following polymers.</strong></p>
    
    <div class="drawing-instructions">
        <strong>Remember:</strong>
        <ul>
            <li>Draw the repeating unit inside brackets [ ]</li>
            <li>Put an 'n' outside the bracket to show it repeats</li>
            <li>Show all bonds clearly</li>
        </ul>
    </div>
    
    <div class="drawing-box">
        <h4>1. Draw the repeating unit of poly(propene)</h4>
        <div class="drawing-space"></div>
    </div>
    
    <div class="drawing-box">
        <h4>2. Draw the repeating unit of poly(chloroethene) / PVC</h4>
        <div class="drawing-space"></div>
    </div>
"""
    
    def generate_data_handling_part(self) -> str:
        """Generate data handling/graphing activity"""
        return """
    <h4>Part 2 – Data Handling</h4>
    
    <p><strong>Plot the data from the table on a bar chart below.</strong></p>
    <div class="hint-box">Remember to: label axes, include units, use a ruler, add a title</div>
    <div class="chart-space"></div>
"""
    
    def generate_matching_activity(self, knowledge_points: list) -> str:
        """Generate matching activity"""
        terms = knowledge_points[:4] if len(knowledge_points) >= 4 else ["Term 1", "Term 2", "Term 3", "Term 4"]
        
        rows_html = ""
        for term in terms:
            rows_html += f"""
        <tr>
            <td class="term">{term[:50]}...</td>
            <td style="width: 30%;"></td>
            <td>Description</td>
        </tr>"""
        
        return f"""
    <h4>Part 2 – Connect the Lines</h4>
    <p>Draw lines to match each <strong>term</strong> to the correct <strong>description</strong>.</p>
    
    <table class="matching-table">
{rows_html}
    </table>
"""
    
    def generate_calculations_part(self, skills_list: list) -> str:
        """Generate calculations based on skills"""
        calc_questions = [s for s in skills_list if 'calculate' in s.lower() or 'work out' in s.lower()]
        
        calc_html = ""
        for i, q in enumerate(calc_questions[:2], 1):
            calc_html += f"""
    <div class="calculation-box">
        <p><strong>{i}. {q}</strong></p>
        <p style="margin-top: 10px;">Show your working:</p>
        <div class="answer-space"></div>
        <p>Answer: __________</p>
    </div>
"""
        
        return f"""
    <h4>Part 3 – Calculations</h4>
{calc_html}
"""
    
    def generate_comparison_table(self, title: str) -> str:
        """Generate comparison table"""
        return f"""
    <h4>Part 3 – Comparison Table</h4>
    <p>Complete the table to compare different aspects of {title}:</p>
    
    <table class="comparison-table">
        <tr>
            <th>Property</th>
            <th>Type A</th>
            <th>Type B</th>
        </tr>
        <tr><td>Structure</td><td></td><td></td></tr>
        <tr><td>Properties</td><td></td><td></td></tr>
        <tr><td>Uses</td><td></td><td></td></tr>
    </table>
"""
    
    def generate_labelling_activity(self) -> str:
        """Generate standard labelling diagram"""
        return """
    <h4>Part 3 – Labelling</h4>
    <div class="diagram-box">
        <div class="diagram-placeholder">
            <p style="text-align: center; color: #999; padding: 80px;">Diagram to label</p>
        </div>
        <p style="text-align: center; margin-top: 10px;"><strong>Label the diagram above</strong></p>
    </div>
"""
    
    def generate_section_b(self, skills_list: list) -> str:
        """Generate Section B questions from DO column"""
        section_b = ""
        
        # Use actual skills from DO column
        for i, skill in enumerate(skills_list[:8], 1):
            # Determine if it's a simple or complex question
            is_explanation = any(word in skill.lower() for word in ['explain', 'describe', 'compare'])
            answer_class = "answer-space-long" if is_explanation else "answer-space"
            
            section_b += f"""
    <p><strong>{i}. {skill}</strong></p>
    <div class="{answer_class}"></div>
"""
        
        return section_b
    
    def extract_keywords(self, text: str) -> list:
        """Extract important keywords from knowledge text"""
        # Simple keyword extraction
        words = re.findall(r'\b[a-z]{4,}\b', text.lower())
        common_words = {'that', 'this', 'with', 'from', 'have', 'they', 'been', 'which', 'their', 'when'}
        keywords = [w for w in words if w not in common_words]
        return list(dict.fromkeys(keywords))[:12]  # Remove duplicates, limit to 12
    
    def generate_worksheet(self, lesson_code: str) -> str:
        """Generate complete HTML worksheet for a lesson"""
        lesson = self.get_lesson(lesson_code)
        if not lesson:
            return None
        
        print(f"\n🔄 Generating worksheet for: {lesson['title']}")
        
        # Analyze what skills are needed
        analysis = self.analyze_skills(lesson['skills'])
        print(f"   📊 Analysis: {analysis}")
        
        # Generate sections
        section_a = self.generate_section_a(lesson, analysis)
        
        # Section B: Use actual skills from DO column
        skills_list = [s.strip() for s in lesson['skills'].split('\n') if s.strip()]
        section_b = self.generate_section_b(skills_list)
        
        # Section C: Extended response
        section_c = f"""
    <p><strong>1. Explain {lesson['title'].lower()} in detail. Include diagrams and examples in your answer. (5-6 marks)</strong></p>
    <div class="answer-space-long"></div>
    <div class="answer-space-long"></div>
    <div class="answer-space-long"></div>
    
    <p><strong>2. A student makes a statement about {lesson['title'].lower()}. Evaluate this statement and explain whether you agree or disagree. (4-5 marks)</strong></p>
    <div class="answer-space-long"></div>
    <div class="answer-space-long"></div>
"""
        
        # HTML template with all styling
        html = self.generate_html_template(lesson, section_a, section_b, section_c)
        
        return html
    
    def generate_html_template(self, lesson: dict, section_a: str, section_b: str, section_c: str) -> str:
        """Generate complete HTML with styling"""
        return f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{lesson['title']} - Worksheet</title>
    <style>
        body {{
            font-family: 'Calibri', Arial, sans-serif;
            max-width: 210mm;
            margin: 0 auto;
            padding: 20px;
            line-height: 1.4;
        }}
        h1 {{ text-align: center; margin-bottom: 30px; }}
        .section-header {{
            background-color: #e0e0e0;
            padding: 10px;
            font-weight: bold;
            font-size: 16px;
            margin-top: 30px;
            margin-bottom: 20px;
            border-left: 5px solid #666;
        }}
        .word-bank {{
            background-color: #f5f5f5;
            padding: 10px;
            margin: 10px 0;
            border: 1px solid #ccc;
        }}
        .fill-blank {{
            list-style-type: decimal;
            margin-left: 20px;
        }}
        .fill-blank li {{ margin: 10px 0; }}
        .matching-table {{
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
        }}
        .matching-table td {{
            padding: 15px;
            border: 1px solid #000;
            vertical-align: middle;
        }}
        .matching-table .term {{
            font-weight: bold;
            width: 30%;
        }}
        .drawing-box {{
            margin: 20px 0;
            padding: 15px;
            border: 2px solid #ccc;
            background-color: #fafafa;
        }}
        .drawing-space {{
            width: 100%;
            height: 200px;
            border: 1px dashed #999;
            background-color: white;
            margin-top: 10px;
        }}
        .drawing-instructions {{
            background-color: #e8f4f8;
            padding: 15px;
            margin: 15px 0;
            border-left: 4px solid #667eea;
        }}
        .drawing-instructions ul {{
            margin-left: 20px;
            margin-top: 10px;
        }}
        .calculation-box {{
            background-color: #f0f8ff;
            padding: 15px;
            margin: 15px 0;
            border: 2px solid #4a90e2;
            border-radius: 5px;
        }}
        .comparison-table {{
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
        }}
        .comparison-table th,
        .comparison-table td {{
            border: 1px solid #000;
            padding: 12px;
            text-align: left;
        }}
        .comparison-table th {{
            background-color: #e0e0e0;
            font-weight: bold;
        }}
        .diagram-box {{
            margin: 20px 0;
            padding: 20px;
            border: 2px solid #ccc;
            background-color: #fafafa;
        }}
        .diagram-placeholder {{
            width: 100%;
            height: 200px;
            border: 2px dashed #999;
            background-color: white;
        }}
        .chart-space {{
            width: 100%;
            height: 300px;
            border: 2px dashed #999;
            background-color: white;
            margin: 20px 0;
        }}
        .answer-space {{
            border-bottom: 1px solid #000;
            min-height: 60px;
            margin: 10px 0;
        }}
        .answer-space-long {{
            border-bottom: 1px solid #000;
            min-height: 100px;
            margin: 10px 0;
        }}
        .hint-box {{
            background-color: #fffacd;
            padding: 10px;
            margin: 10px 0;
            border-left: 4px solid #ffd700;
            font-style: italic;
            font-size: 14px;
        }}
    </style>
</head>
<body>
    <h1>{lesson['title']} - {lesson['code']}</h1>
    <p><strong>Name:</strong> _______________________________  <strong>Date:</strong> _______________</p>
    
    <div class="section-header">Section A – Knowledge Recall (Support Section)</div>
{section_a}
    
    <div class="section-header">Section B – Understanding (Main Section)</div>
{section_b}
    
    <div class="section-header">Section C – Extended Response</div>
{section_c}

</body>
</html>"""
    
    def save_worksheet(self, html: str, lesson_code: str, output_dir: str = "worksheets"):
        """Save HTML worksheet to file"""
        Path(output_dir).mkdir(exist_ok=True)
        filename = f"{lesson_code}_worksheet.html"
        filepath = Path(output_dir) / filename
        filepath.write_text(html, encoding='utf-8')
        print(f"✅ Saved: {filename}")
        return filepath
    
    def save_word_doc(self, html: str, lesson_code: str, output_dir: str = "worksheets"):
        """Save worksheet as Word document"""
        doc = Document()
        text = re.sub('<[^<]+?>', '', html)  # Strip HTML tags
        doc.add_paragraph(text)
        filename = f"{lesson_code}_worksheet.docx"
        doc.save(Path(output_dir) / filename)
        print(f"✅ Saved: {filename}")


def main():
    """Main execution"""
    # Configuration
    EXCEL_FILE = "C4.1 Unit Overview.xlsx"  # Change this to your file path
    
    # Initialize generator
    generator = WorksheetGenerator(EXCEL_FILE)
    
    # Generate worksheet for specific lesson
    lesson_code = "C4.1.12"  # Change this to any lesson code
    
    html = generator.generate_worksheet(lesson_code)
    if html:
        generator.save_word_doc(html, lesson_code)
        print(f"\n✨ Done! Word document created in 'worksheets/{lesson_code}_worksheet.docx'")
    
    # To generate all lessons:
    # for code in ["C4.1.1", "C4.1.2", "C4.1.3", ...]:
    #     html = generator.generate_worksheet(code)
    #     if html:
    #         generator.save_word_doc(html, code)


if __name__ == "__main__":
    main()