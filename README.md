# Education Resource Generator

A collection of Python tools for generating educational materials including PowerPoint slides, worksheets, ReadNows, and Think Out Loud scripts.

## Project Structure

```
education/
├── Slides/                  # PowerPoint slide generator
│   ├── generate.py         # Main generator script
│   ├── config.py           # Configuration (target lesson, units, etc.)
│   ├── lib/                # Core processing modules
│   ├── templates/          # PowerPoint templates
│   └── output/             # Generated PowerPoint files
│
├── readnow/                 # ReadNow document generator
│   ├── unified_generator.py # Main generator
│   ├── main.py             # Core functions
│   ├── constants.py        # Configuration (year, reading age, etc.)
│   └── readnows/           # Generated ReadNow documents
│
├── worksheet/               # Worksheet generator
│   └── worksheets/          # Generated worksheets
│
├── thinkOutLoud/            # Think Out Loud script generator
│   └── *.docx              # Generated scripts
│
├── Lesson Resources/        # Source materials (PowerPoints, Excel files)
└── pyproject.toml          # Python dependencies (uses uv)
```

## Quick Start

### Daily Review / Core Knowledge (for colleagues)

The **Daily Review** app generates AQA-style Core Knowledge .docx (weekly words + “State…” questions). Colleagues can use it **without coding**:

- **One person** runs the app (e.g. double‑click **Run Daily Review app.command** on Mac or **Run Daily Review app.bat** on Windows), then shares the **Network URL** (e.g. `http://192.168.x.x:8501`) with everyone.
- **Everyone else** opens that link in their browser and uses the app like a normal website.

Full steps (and an optional “put it online” option): **[SHARING_WITH_COLLEAGUES.md](SHARING_WITH_COLLEAGUES.md)**.

### Web Interface (Simple Frontend)

The easiest way to use the generators:

```bash
uv run streamlit run app.py
```

This opens a simple web interface in your browser with buttons to generate each resource type.

### Generate PowerPoint Slides

```bash
cd Slides
uv run python generate.py
```

Configure in `Slides/config.py`:
- `TARGET_LESSON` - Specific lesson code (e.g., "B3.2.4") or `None` for all
- `TARGET_UNIT` - Specific unit (e.g., "B3.2") or `None` for all

### Generate ReadNow Documents

```bash
cd readnow
uv run python unified_generator.py
```

Configure in `readnow/constants.py`:
- `STUDENT_YEAR` - Year group (e.g., "Year 9")
- `STUDENT_ATTAINMENT` - "HPA" or "LPA"
- `READING_AGE` - Target reading age
- `WORD_COUNT` - Word count range (e.g., "100-120")
- `LESSON_CODE_PREFIX` - Filter lessons (e.g., "C3.2.")

### Generate Worksheets

```bash
cd worksheet
uv run python main.py
```

## Features

### PowerPoint Generator (`Slides/`)
- Extracts objectives from source PowerPoints (slides 4-7)
- Retrieves markschemes from ReadNow documents
- Retrieves exit tickets from Excel spreadsheets
- Uses templates for consistent formatting
- Auto-detects lesson data from Excel files

### ReadNow Generator (`readnow/`)
- AI-generated content based on lesson objectives
- Supports HPA and LPA variants
- Customizable reading age and word count
- Searches multiple slides (4-7) to find objectives

### Worksheet Generator (`worksheet/`)
- AI-generated practice questions
- Multiple question types (recall, understanding, extended)
- Automatic mark scheme generation

## Dependencies

Install using `uv`:

```bash
uv sync
```

Key dependencies:
- `python-pptx` - PowerPoint manipulation
- `python-docx` - Word document generation
- `pandas` - Excel file processing
- `openai` / `anthropic` - AI content generation

## Data Sources

- **Excel Files**: `Lesson Resources/**/Unit Guidance/**/*.xlsx`
  - Contains lesson codes, titles, objectives, exit tickets
- **PowerPoint Files**: `Lesson Resources/**/*.pptx`
  - Source slides for extracting objectives
- **ReadNow Documents**: `readnow/readnows/*.docx`
  - Source for markschemes

## Configuration

Each module has its own config file:
- `Slides/config.py` - Slide generation settings
- `readnow/constants.py` - ReadNow generation settings
# education
