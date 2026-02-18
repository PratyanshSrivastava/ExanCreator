# 📝 Exam Paper Generator – NCERT India (Flask)

A professional web application for generating print-ready examination papers as `.docx` files for Indian schools (Class 6–12).

## Quick Start

### Requirements
- Python 3.8+
- pip

### Installation

```bash
pip install flask python-docx
```

### Run

```bash
cd exam_app
python app.py
```

Then open: **http://localhost:5000**

---

## Features

### ✅ Paper Metadata
- School name, address
- Class (6–12), Subject (30+ NCERT subjects with auto-codes)
- Exam type (Annual, Unit Test, Pre-Board, etc.)
- Set (A/B/C), Marks, Duration, Academic Year

### ✅ Section Management
- Add unlimited sections (A, B, C…)
- Per-section name, marks, instructions, question numbering style
- Collapse/expand sections

### ✅ Question Types (15 types)
1. **MCQ** – Multiple choice with A/B/C/D options
2. **Very Short Answer (VSA)**
3. **Short Answer (SA)**
4. **Long Answer (LA)** with answer lines
5. **Fill in the Blanks**
6. **Match the Following**
7. **True / False**
8. **Assertion–Reason**
9. **Comprehension / Unseen Passage**
10. **Case-Based Questions**
11. **Numerical Problems**
12. **Definition / One-word Answer**
- Each question supports **Parts (a), (b), (c)…**

### ✅ Professional .docx Output
- A4 paper, 1" margins, Times New Roman 12pt
- School header with name and exam details
- Numbered instructions block
- 3-column question table (Q.No | Question | Marks)
- Answer lines for SA/LA questions
- Page numbers in footer

### ✅ Live Preview
- In-browser preview before downloading
- Approximate replica of final output

### ✅ Save / Load Drafts
- Auto-save to localStorage
- `Ctrl+S` to save, load from sidebar

---

## Project Structure

```
exam_app/
├── app.py              # Flask routes
├── docx_generator.py   # .docx generation (python-docx)
└── templates/
    └── index.html      # Complete frontend (HTML/CSS/JS)
```

## API Endpoints

| Method | URL | Description |
|--------|-----|-------------|
| GET | `/` | Web UI |
| GET | `/api/subjects` | List of subjects |
| POST | `/api/paper/generate` | Generate .docx (returns file) |

### POST /api/paper/generate

Send JSON body:
```json
{
  "metadata": { ... },
  "instructions": ["..."],
  "sections": [{ "questions": [...] }]
}
```
Returns a `.docx` file download.
