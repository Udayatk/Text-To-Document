from flask import Flask, request, jsonify, send_file
from doc_builder import save_markdown, save_pdf, save_word, save_excel
import os
from werkzeug.utils import secure_filename
import csv
from io import StringIO

app = Flask(__name__)
chat_history = []

@app.route('/upload', methods=['POST'])
def upload():
    # Accept either file upload or JSON rows
    if request.is_json:
        data = request.get_json()
        rows = data.get('rows')
        if not rows:
            return jsonify({'error': 'No data provided'}), 400
        chat_history.clear()
        chat_history.extend(rows)
        return jsonify({'message': 'CSV data processed and chat history updated.'})
    elif 'file' in request.files:
        file = request.files['file']
        if not file or not file.filename:
            return jsonify({'error': 'No selected file'}), 400
        filename = secure_filename(file.filename or "uploaded.txt")
        text = file.read().decode('utf-8')
        reader = csv.DictReader(StringIO(text))
        chat_history.clear()
        chat_history.extend([row for row in reader])
        return jsonify({'message': 'File processed and chat history updated.'})
    else:
        return jsonify({'error': 'No file part or JSON data'}), 400

@app.route('/')
def home():
    return '''<h2>Chatbot to Document Builder</h2>
    <p>Use <b>POST /chat</b> with JSON {"message": "your text"} to chat.<br>
    Use <b>GET /export</b> to download the Markdown documentation.</p>'''

@app.route('/chat', methods=['POST'])
def chat():
    if not request.is_json:
        return jsonify({'error': 'Request must be JSON with Content-Type: application/json'}), 400
    data = request.get_json()
    if not data or 'message' not in data:
        return jsonify({'error': 'Missing "message" in request body'}), 400
    user_message = data['message']
    chat_history.append({'message': user_message})
    return jsonify({'reply': ''})

@app.route('/export', methods=['GET'])
def export():
    if not chat_history or not isinstance(chat_history, list) or not all(isinstance(row, dict) for row in chat_history):
        error_detail = {
            'chat_history_type': str(type(chat_history)),
            'chat_history_length': len(chat_history) if isinstance(chat_history, list) else 'N/A',
            'chat_history_sample': chat_history[:2] if isinstance(chat_history, list) else str(chat_history)
        }
        return jsonify({'error': 'No chat history to export or invalid format.', 'detail': error_detail}), 400
    doc_type = request.args.get('type', 'md')
    # Normalize chat_history for single-column export
    normalized_history = chat_history
    if isinstance(chat_history, list) and len(chat_history) > 0:
        # If all rows have only 'user' key, convert to 'message' for compatibility
        if all(list(row.keys()) == ['user'] for row in chat_history):
            normalized_history = [{'message': row['user']} for row in chat_history]
    try:
        if doc_type == 'pdf':
            doc_path = save_pdf(normalized_history)
            mime = 'application/pdf'
        elif doc_type == 'word':
            doc_path = save_word(normalized_history)
            mime = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        elif doc_type == 'excel':
            doc_path = save_excel(normalized_history)
            mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        else:
            doc_path = save_markdown(normalized_history)
            mime = 'text/markdown'
        return send_file(doc_path, as_attachment=True, mimetype=mime)
    except Exception as e:
        return jsonify({'error': f'Failed to export: {str(e)}'}), 500

if __name__ == '__main__':
    app.run(debug=True)