# 🎓 Google Gemini Assignment Solver Agent

An automated, intelligent batch agent powered by **Google Gemini AI** and **Python-Docx** to solve and format `.docx` assignment worksheets in bulk.

## 🚀 Features

- **Batch Processing**: Processes 18-20+ `.docx` files in seconds with progress tracking and summary reports.
- **Intelligent Gemini AI**: Dynamically extracts course details, topics, and questions to query Google Gemini for accurate, syllabus-aligned answers.
- **Student Metadata Injection**: Automatically injects student credentials across all worksheets:
  - **Name**: `Sanchit Nigam`
  - **Registration No**: `RA2311003030178`
  - Roll-number fields use the configured registration number.
  - **Date**: Formatted current date (`DD/MM/YYYY`) or customizable.
- **Formatting Preservation**: Strictly preserves *Times New Roman* typography, font sizes, margins, alignments, headings, logos/images, and footers (`Dr.Divya Mohan ... 21CSE458T`).
- **Answer Highlighting**: Fills blanks with bold and underlined text matching the document's design.
- **Open-Ended Answers**: Appends answers beneath an `ANSWERS` heading without duplicating an existing answer heading; responses use paragraphs or bullets, never tables.
- **Rate-Limit & Error Handling**: Exponential backoff retry logic to handle high-volume bulk processing.
- **Dry-Run & Mock Modes**: Test without modifying files or consuming API credits.

---

## 📦 Project Structure

```
assignment-agent/
├── InputFolder/          # Directory containing input .docx files (e.g. 1011.docx, 1012.docx...)
├── OutputFolder/         # Directory where solved .docx files are saved
├── config.py             # Configuration parameters & defaults
├── docx_parser.py        # Extracts topics, footers, student metadata, & questions
├── gemini_solver.py      # Google GenAI API client with structured JSON prompting
├── docx_writer.py        # Injects metadata & answers with preserved styles
├── main.py               # CLI entry point & Rich batch runner
├── .env.example          # Template for Gemini API key
├── pyproject.toml        # Dependencies & package metadata
└── README.md
```

---

## 🛠️ Setup & Installation

1. **Clone or Navigate to the Workspace**:
   ```bash
   cd /home/sanchitnigam/Desktop/Coding/Python/General/assignment-agent
   ```

2. **Install Dependencies** (using `uv` or `pip`):
   ```bash
   uv sync
   # or: pip install -e .
   ```

3. **Set Up Gemini API Key**:
   Create a `.env` file from the template:
   ```bash
   cp .env.example .env
   ```
   Add your Gemini API key inside `.env`:
   ```ini
   GEMINI_API_KEY=AIzaSy...
   ```

---

## 💻 Usage

### 1. Batch Solve All Files with Gemini
```bash
uv run python main.py
```
This reads all `.docx` files in `InputFolder/` and saves the completed worksheets into `OutputFolder/`.

### 2. Overwrite Files In-Place
```bash
uv run python main.py --in-place
```

### 3. Custom Input/Output Directories
```bash
uv run python main.py --input-dir /path/to/my_inputs --output-dir /path/to/my_outputs
```

### 4. Provide API Key Directly via CLI
```bash
uv run python main.py --api-key YOUR_GEMINI_API_KEY
```

### 5. Dry-Run Mode (Preview without saving)
```bash
uv run python main.py --dry-run
```

### 6. Offline Mock Mode (For testing)
```bash
uv run python main.py --mock
```

---

## ⚙️ CLI Options Reference

| Flag | Default | Description |
|---|---|---|
| `--input-dir`, `-i` | `InputFolder` | Folder containing source `.docx` worksheets |
| `--output-dir`, `-o` | `OutputFolder` | Folder to save updated `.docx` files |
| `--in-place` | `False` | Overwrite input files directly |
| `--api-key` | Env `GEMINI_API_KEY` | Google Gemini API key |
| `--model` | `gemini-2.5-flash` | Gemini model to use (`gemini-2.5-flash`, `gemini-1.5-flash`, etc.) |
| `--name` | `Sanchit Nigam` | Student name to inject |
| `--reg-no` | `RA2311003030178` | Registration number to inject |
| `--date` | Current Date (`DD/MM/YYYY`) | Date string to inject |
| `--dry-run` | `False` | Preview solutions in terminal without writing |
| `--mock` | `False` | Offline mock solver for testing |
