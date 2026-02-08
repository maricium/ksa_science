"""Read data from Excel and ReadNow files"""

import pandas as pd
from docx import Document
from .utils import find_path


def read_lesson_data(excel_file, sheet_name):
    """Read lesson data from Excel spreadsheet"""
    try:
        # First read raw to check header row for exit ticket column and detect by content
        df_raw = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
        exit_ticket_col_idx = None
        
        # Check row 3 (header row) for exit ticket column
        if len(df_raw) > 3:
            for col_idx in range(len(df_raw.columns)):
                header_val = df_raw.iloc[3, col_idx]
                if pd.notna(header_val) and 'exit' in str(header_val).lower() and 'ticket' in str(header_val).lower():
                    exit_ticket_col_idx = col_idx
                    break
        
        # If not found by header, try detecting by content pattern (multiple choice questions)
        if exit_ticket_col_idx is None:
            for col_idx in range(len(df_raw.columns)):
                # Check data rows (4+) for exit ticket pattern (multiple choice: A. B. C.)
                for row_idx in range(4, min(len(df_raw), 15)):
                    cell_val = df_raw.iloc[row_idx, col_idx]
                    if pd.notna(cell_val):
                        val_str = str(cell_val).strip()
                        # Look for multiple choice pattern (A. B. C. or A) B) C))
                        if (len(val_str) > 50 and 
                            ('\nA.' in val_str or '\nA)' in val_str or ' A. ' in val_str) and
                            ('B.' in val_str or 'B)' in val_str) and
                            ('C.' in val_str or 'C)' in val_str)):
                            exit_ticket_col_idx = col_idx
                            break
                if exit_ticket_col_idx is not None:
                    break
        
        # Now read with header=3 for proper structure
        df = pd.read_excel(excel_file, sheet_name=sheet_name, header=3).dropna(how='all')
        lesson_code_col = next((i for i in range(min(3, len(df.columns))) if df.iloc[:, i].astype(str).str.contains(r'[CPB]\d+\.\d+\.\d+', na=False, regex=True).any()), 1)
        
        mapping = {df.columns[lesson_code_col]: 'lesson_code'}
        if len(df.columns) > lesson_code_col + 1:
            mapping[df.columns[lesson_code_col + 1]] = 'lesson_title'
        if len(df.columns) > lesson_code_col + 2:
            mapping[df.columns[lesson_code_col + 2]] = 'knowledge_objectives'
        
        # Map exit ticket column if we found it
        if exit_ticket_col_idx is not None and exit_ticket_col_idx < len(df.columns):
            # Make sure we don't overwrite an existing mapping
            col_name = df.columns[exit_ticket_col_idx]
            if col_name not in mapping.values():
                mapping[col_name] = 'exit_ticket'
        
        for col_idx, col_name in enumerate(df.columns):
            if col_name not in mapping.values():
                sample = df.iloc[:5, col_idx].astype(str).str.cat(sep=' ').lower()
                col_name_lower = str(col_name).lower()
                if 'skill' in col_name_lower or ('skill' in sample or 'do' in sample):
                    # Only map to skill_objectives if not already mapped
                    if col_name not in mapping:
                        mapping[col_name] = 'skill_objectives'
                elif ('exit' in col_name_lower and 'ticket' in col_name_lower):
                    mapping[col_name] = 'exit_ticket'
        
        df = df.rename(columns=mapping)
        if 'lesson_code' in df.columns:
            df = df[df['lesson_code'].astype(str).str.contains(r'[CPB]\d+\.\d+\.\d+', na=False, regex=True)]
        return df
    except Exception as e:
        print(f"Error reading Excel: {e}")
        # Try fallback engines if primary fails
        for fallback_engine in ['openpyxl', 'xlrd']:
            try:
                return pd.read_excel(excel_file, sheet_name=sheet_name, header=0, engine=fallback_engine).dropna(how='all')
            except:
                continue
        raise


def read_markscheme(lesson_code):
    """Read MARK SCHEME from ReadNow documents"""
    readnows_dir = find_path("readnow/readnows")
    if not readnows_dir.exists():
        return "MARK SCHEME\n\nMarkscheme not available."
    
    for pattern in [f"{lesson_code}_ReadNow_LPA.docx", f"{lesson_code}_ReadNow_HPA.docx", f"{lesson_code}_ReadNow.docx"]:
        found = list(readnows_dir.rglob(pattern))
        if found:
            try:
                doc = Document(str(found[0]))
                content, found_ms = [], False
                for para in doc.paragraphs:
                    if para.text.strip():
                        if "MARK SCHEME" in para.text.upper():
                            found_ms = True
                            content.append(para.text.strip())
                        elif found_ms:
                            if para.text.isupper() and len(para.text) > 10:
                                break
                            content.append(para.text.strip())
                if content:
                    print(f"Found MARK SCHEME for {lesson_code}")
                    return _format_markscheme('\n'.join(content))
            except Exception as e:
                print(f"Error reading ReadNow: {e}")
    return "MARK SCHEME\n\nMarkscheme not available."


def _format_markscheme(content):
    """Format markscheme content"""
    lines, q_num = [], 1
    for line in content.split('\n'):
        line = line.strip()
        if not line:
            continue
        if "MARK SCHEME" in line.upper():
            lines.append(line)
        elif line.startswith(('•', '-', '*')):
            lines.append(f"{q_num}. {line.lstrip('•-* ').strip()}")
            q_num += 1
        elif line and not line.startswith(tuple(f"{i}." for i in range(1, 20))):
            lines.append(f"{q_num}. {line}")
            q_num += 1
        else:
            lines.append(line)
    return '\n'.join(lines)
