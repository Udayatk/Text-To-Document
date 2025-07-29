# Project Repository: https://github.com/Udayatk/Text-To-Document-Markdown-Pdf-Word-Excell

# Chat History Document Builder

A Python Flask backend and Streamlit frontend for converting chat history (CSV or manual input) into professional documents: Markdown, PDF, Word, and Excel.

---

## Features
- **Upload chat history** via CSV file or manual entry
- **Export** in Markdown, PDF, Word, or Excel formats
- **Professional formatting**: Wikipedia-style paragraphs, bold headings, links, and clean tables
- **Streamlit UI** for easy interaction
- **REST API** for integration with other tools

---

## How It Works
1. **Start the Flask backend**
   - `python app.py` (or use your virtual environment)
2. **Start the Streamlit frontend**
   - `streamlit run streamlit_app.py`
3. **Upload chat history**
   - Use the UI to upload a CSV or enter text manually
4. **Export your document**
   - Choose your format and download the file

---

## API Endpoints
- `POST /upload` — Upload chat history (CSV or JSON)
- `GET /export?type=md|pdf|word|excel` — Download the document in your chosen format

---

## Formatting Logic
- **Single-column:** Each message is a clean, spaced paragraph
- **Multi-column:** Data is exported as a table in all formats
- **PDF/Word:** Headings, bold phrases, links, and spacing styled for professional output

---

## Example Input
```markdown
| First Name | Last Name | Department | Supervisor |
|------------|-----------|------------|------------------------|
| Fareed     | Awad      | Marketing  | Jane Doe, Vice-President |
| Adam       | Doe       | Custodial  | Donna Martin, C.E.O.     |
| Jane       | Doe       | Executive  | Donna Martin, C.E.O.     |
| Donna      | Martin    | Executive  | None                     |
| John       | Smith     | Marketing  | Jane Doe, Vice-President |
```

Paste a Markdown table like the above in the UI and select "Word" for export. The output will be a clean, professional table in your Word document, with no header/separator rows and no extra spaces.

---

## Requirements
- Python 3.11+
- Flask
- Streamlit
- pandas, fpdf, python-docx

---

## Quick Start
```bash
pip install -r requirements.txt
python app.py
streamlit run streamlit_app.py
```

---

## License
MIT

---

## Author
Udaya (and GitHub Copilot)
