# Text to Document Builder

A Streamlit app to convert chat/text history into Markdown, PDF, Word, and Excel documents.

## Features
- Manual text input or file upload
- Export to Markdown, PDF, Word (.docx), Excel (.xlsx)
- Markdown table detection for Word/Excel
- Download generated documents

## How to Run
1. Install dependencies:
   ```
   pip install -r requirements.txt
   ```
2. Start the app:
   ```
   streamlit run text_to_document.py
   ```

## Usage
- Select document type (Markdown, PDF, Word, Excel)
- Enter text or upload a file
- Click 'Process' then 'Download' to get your document

## File Structure
- `text_to_document.py`: Main Streamlit app
- `requirements.txt`: Python dependencies
- `README.md`: Project info and instructions

Deployed - https://text-to-document.streamlit.app/
