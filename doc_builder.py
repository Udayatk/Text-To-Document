import pandas as pd
from fpdf import FPDF
from docx import Document
import datetime
import os

def get_output_path(filename):
    return os.path.join(os.getcwd(), filename)

def save_excel(chat_history):
    filename = f"chat_document_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    out_path = get_output_path(filename)
    # Expect chat_history as list of dicts with keys from CSV header
    df = pd.DataFrame(chat_history)
    df.to_excel(out_path, index=False)
    return out_path

def save_markdown(chat_history):
    filename = f"chat_document_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.md"
    out_path = get_output_path(filename)
    if chat_history:
        header = list(chat_history[0].keys())
        with open(out_path, 'w', encoding='utf-8') as f:
            if len(header) == 1:
                # Single column, write each message as a paragraph (no bullets)
                for entry in chat_history:
                    msg = str(entry[header[0]]).strip()
                    if msg:
                        f.write(msg + '\n\n')
            else:
                f.write('| ' + ' | '.join(header) + ' |\n')
                f.write('| ' + ' | '.join(['---']*len(header)) + ' |\n')
                for entry in chat_history:
                    f.write('| ' + ' | '.join(str(entry[k]) for k in header) + ' |\n')
    return out_path

def save_pdf(chat_history):
    import unicodedata
    filename = f"chat_document_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    out_path = get_output_path(filename)
    try:
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=10)
        if chat_history:
            header = list(chat_history[0].keys())
            if len(header) == 1:
                # Enhanced format: headings, bold phrases, links, Unicode, clean spacing
                import re
                valid = False
                for entry in chat_history:
                    msg = str(entry[header[0]]).strip()
                    if not msg:
                        continue
                    valid = True
                    # Unicode support: no forced ASCII
                    safe_text = msg[:1000]
                    # Heading: all-caps or ends with colon
                    if safe_text.isupper() and len(safe_text) > 2 or safe_text.endswith(":"):
                        pdf.set_font("Arial", style="B", size=14)
                        pdf.multi_cell(0, 12, safe_text)
                        pdf.set_font("Arial", size=11)
                        pdf.ln(6)
                    # Bold phrase before colon or parentheses
                    elif (":" in safe_text or "(" in safe_text) and not safe_text.isupper():
                        match = re.match(r"^([^.\n:()]+[:(][^\n]*)", safe_text)
                        if match:
                            bold_phrase = match.group(1)
                            rest = safe_text[len(bold_phrase):].lstrip()
                            pdf.set_font("Arial", style="B", size=11)
                            pdf.write(8, bold_phrase)
                            pdf.set_font("Arial", size=11)
                            pdf.write(8, " " + rest + "\n")
                            pdf.ln(4)
                        else:
                            pdf.set_font("Arial", size=11)
                            pdf.multi_cell(0, 8, safe_text)
                            pdf.ln(4)
                    # Link: underline and blue
                    elif re.search(r"https?://\S+", safe_text):
                        pdf.set_text_color(0, 0, 255)
                        pdf.set_font("Arial", style="U", size=11)
                        pdf.multi_cell(0, 8, safe_text)
                        pdf.set_text_color(0, 0, 0)
                        pdf.set_font("Arial", size=11)
                        pdf.ln(4)
                    else:
                        pdf.set_font("Arial", size=11)
                        pdf.multi_cell(0, 8, safe_text)
                        pdf.ln(4)
                if not valid:
                    pdf.set_font("Arial", size=11)
                    pdf.multi_cell(0, 8, "(No valid messages provided)")
            else:
                # Table format
                col_widths = []
                for h in header:
                    max_len = max([len(str(row[h])) for row in chat_history] + [len(h)])
                    col_widths.append(max(25, min(60, max_len * 3)))  # min/max width per column
                row_height = pdf.font_size + 4
                # Header
                for i, h in enumerate(header):
                    pdf.cell(col_widths[i], row_height, h, border=1, align='C')
                pdf.ln(row_height)
                # Rows
                for entry in chat_history:
                    y_before = pdf.get_y()
                    x_start = pdf.get_x()
                    max_cell_height = row_height
                    cell_texts = [str(entry[k]) for k in header]
                    # Calculate max cell height for this row
                    cell_heights = []
                    for i, text in enumerate(cell_texts):
                        safe_text = unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('ascii')
                        cell_height = pdf.get_string_width(safe_text) // col_widths[i] * row_height + row_height
                        cell_heights.append(cell_height)
                    max_cell_height = max(cell_heights)
                    for i, text in enumerate(cell_texts):
                        safe_text = unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('ascii')
                        x = pdf.get_x()
                        y = pdf.get_y()
                        pdf.multi_cell(col_widths[i], row_height, safe_text, border=1, align='L', max_line_height=row_height)
                        pdf.set_xy(x + col_widths[i], y)
                    pdf.ln(max_cell_height)
        pdf.output(out_path)
        return out_path
    except Exception as e:
        raise RuntimeError(f"PDF generation failed: {str(e)}")

def save_word(chat_history):
    filename = f"chat_document_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
    out_path = get_output_path(filename)
    doc = Document()
    import re
    from docx.shared import RGBColor, Pt
    if chat_history:
        header = list(chat_history[0].keys())
        if len(header) == 1:
            for entry in chat_history:
                msg = str(entry[header[0]]).strip()
                if not msg:
                    continue
                # Heading: all-caps or ends with colon
                if (msg.isupper() and len(msg) > 2) or msg.endswith(":"):
                    p = doc.add_paragraph()
                    run = p.add_run(msg)
                    run.bold = True
                    run.font.size = Pt(14)
                    p.paragraph_format.space_after = Pt(12)
                # Bold phrase before colon or parentheses
                elif (":" in msg or "(" in msg) and not msg.isupper():
                    match = re.match(r"^([^.\n:()]+[:(][^\n]*)", msg)
                    if match:
                        bold_phrase = match.group(1)
                        rest = msg[len(bold_phrase):].lstrip()
                        p = doc.add_paragraph()
                        run_bold = p.add_run(bold_phrase)
                        run_bold.bold = True
                        run_bold.font.size = Pt(12)
                        run_rest = p.add_run(" " + rest)
                        run_rest.font.size = Pt(12)
                        p.paragraph_format.space_after = Pt(8)
                    else:
                        p = doc.add_paragraph(msg)
                        p.paragraph_format.space_after = Pt(8)
                # Link: underline and blue
                elif re.search(r"https?://\S+", msg):
                    p = doc.add_paragraph()
                    for part in re.split(r"(https?://\S+)", msg):
                        if re.match(r"https?://\S+", part):
                            run = p.add_run(part)
                            run.font.color.rgb = RGBColor(0, 0, 255)
                            run.underline = True
                        else:
                            p.add_run(part)
                    p.paragraph_format.space_after = Pt(8)
                else:
                    p = doc.add_paragraph(msg)
                    p.paragraph_format.space_after = Pt(8)
        else:
            table = doc.add_table(rows=1, cols=len(header))
            hdr_cells = table.rows[0].cells
            for i, h in enumerate(header):
                hdr_cells[i].text = h
            for entry in chat_history:
                row_cells = table.add_row().cells
                for i, k in enumerate(header):
                    row_cells[i].text = str(entry[k])
    doc.save(out_path)
    return out_path
