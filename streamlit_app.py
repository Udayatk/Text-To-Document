import streamlit as st
import requests
import io
import csv
from io import StringIO

st.title("Text to Document Builder")
st.write("""
Send a message to the chatbot and export the conversation as Markdown documentation.
""")

# 1. Select document type
st.subheader("Step 1: Choose document type")
doc_type = st.selectbox("Select document type to export:", ["Markdown", "PDF", "Word", "Excel"])
type_map = {"Markdown": "md", "PDF": "pdf", "Word": "word", "Excel": "excel"}

# 2. Manual text input
if doc_type == "Excel":
    st.subheader("Step 2: Type CSV text to convert to Excel document")
    typed_text = st.text_area("Enter your CSV text here:")
    if typed_text:
        if st.button("Process Typed Text"):
            reader = csv.DictReader(StringIO(typed_text))
            rows = [row for row in reader]
            if not reader.fieldnames or len(reader.fieldnames) < 1:
                st.error("CSV must have a header row (e.g. name,age)")
            elif not rows or all(all(v == '' for v in row.values()) for row in rows):
                st.error("CSV must have at least one data row below the header.")
            else:
                response = requests.post("http://127.0.0.1:5000/upload", json={"rows": rows})
                if response.status_code == 200:
                    st.success("CSV text processed and chat history updated.")
                else:
                    try:
                        error_msg = response.json().get("error", "Error processing text.")
                    except Exception:
                        error_msg = response.text or "Error processing text."
                    st.error(error_msg)
else:
    st.subheader(f"Step 2: Type text to convert to {doc_type} document")
    typed_text = st.text_area(f"Enter your text for {doc_type} here:")
    if typed_text:
        if st.button("Process Typed Text"):
            # Treat each line as a chat message for non-Excel formats
            rows = []
            for line in typed_text.splitlines():
                if line.strip():
                    rows.append({"message": line.strip()})
            if not rows:
                st.error(f"Please enter at least one line of text for {doc_type}.")
            else:
                response = requests.post("http://127.0.0.1:5000/upload", json={"rows": rows})
                if response.status_code == 200:
                    st.success(f"Text processed and chat history updated for {doc_type}.")
                else:
                    try:
                        error_msg = response.json().get("error", f"Error processing text for {doc_type}.")
                    except Exception:
                        error_msg = response.text or f"Error processing text for {doc_type}."
                    st.error(error_msg)

# 3. File upload
if doc_type == "Excel":
    st.subheader("Step 3: Or upload a CSV file")
    uploaded_file = st.file_uploader("Choose a CSV file to convert:", type=["csv", "txt"])
    if uploaded_file:
        content = uploaded_file.getvalue().decode('utf-8')
        reader = csv.DictReader(StringIO(content))
        rows = [row for row in reader]
        if not reader.fieldnames or len(reader.fieldnames) < 1:
            st.error("CSV file must have a header row (e.g. name,age)")
        elif not rows or all(all(v == '' for v in row.values()) for row in rows):
            st.error("CSV file must have at least one data row below the header.")
        else:
            response = requests.post("http://127.0.0.1:5000/upload", json={"rows": rows})
            if response.status_code == 200:
                st.success("File processed and chat history updated.")
            else:
                try:
                    error_msg = response.json().get("error", "Error processing file.")
                except Exception:
                    error_msg = response.text or "Error processing file."
                st.error(error_msg)
else:
    st.subheader(f"Step 3: Or upload a text file for {doc_type}")
    uploaded_file = st.file_uploader(f"Choose a text file to convert to {doc_type}:", type=["txt", "md", "docx"])
    if uploaded_file:
        content = uploaded_file.getvalue().decode('utf-8')
        # Treat each line as a chat message
        rows = []
        for line in content.splitlines():
            if line.strip():
                rows.append({"message": line.strip()})
        if not rows:
            st.error(f"Text file must have at least one line for {doc_type}.")
        else:
            response = requests.post("http://127.0.0.1:5000/upload", json={"rows": rows})
            if response.status_code == 200:
                st.success(f"File processed and chat history updated for {doc_type}.")
            else:
                try:
                    error_msg = response.json().get("error", f"Error processing file for {doc_type}.")
                except Exception:
                    error_msg = response.text or f"Error processing file for {doc_type}."
                st.error(error_msg)

# 4. Download/export button
st.subheader(f"Step 4: Download your {doc_type} document")
if st.button(f"Download {doc_type}"):
    export_response = requests.get(f"http://127.0.0.1:5000/export?type={type_map[doc_type]}")
    if export_response.status_code == 200:
        file_ext = type_map[doc_type]
        mime_map = {
            "md": "text/markdown",
            "pdf": "application/pdf",
            "word": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "excel": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        }
        st.download_button(
            label=f"Download {doc_type}",
            data=export_response.content,
            file_name=f"chat_document.{file_ext if file_ext != 'excel' else 'xlsx'}",
            mime=mime_map[file_ext]
        )
    else:
        st.error("Failed to export chat history.")

