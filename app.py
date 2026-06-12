import os
import io
import tempfile
from flask import Flask, request, render_template, jsonify, send_file
from werkzeug.utils import secure_filename
from core.extractor import extract_terms, extract_raw_text
from core.formatter import generate_excel_glossary

app = Flask(__name__)

@app.errorhandler(Exception)
def handle_exception(e):
    """Catch all unhandled exceptions to return JSON instead of default HTML error pages"""
    if hasattr(e, 'code') and hasattr(e, 'description'):
        return jsonify({"error": e.description}), e.code
    return jsonify({"error": str(e)}), 500

@app.route('/')
def index():
    return render_template('index.html')

@app.after_request
def add_header(response):
    """Disable caching for all responses to avoid template and JS caching in browsers"""
    response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, post-check=0, pre-check=0, max-age=0'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '-1'
    return response

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'docx_file' not in request.files:
        return jsonify({"error": "No upload file found"}), 400

    docx_file = request.files['docx_file']
    if docx_file.filename == '':
        return jsonify({"error": "No selected file"}), 400

    filename = secure_filename(docx_file.filename)
    allowed_extensions = {'.docx', '.pdf'}
    _, ext = os.path.splitext(filename.lower())
    if ext not in allowed_extensions:
        return jsonify({"error": "Only .docx or .pdf files are allowed"}), 400

    # Use tempfile to write temporary file without persisting to disk (Vercel-friendly)
    fd, temp_path = tempfile.mkstemp(suffix=ext)
    try:
        with os.fdopen(fd, 'wb') as tmp:
            docx_file.save(tmp)
        
        # Extract keywords directly (synchronous)
        core_terms = extract_terms(temp_path)
        raw_text = extract_raw_text(temp_path)
        
        return jsonify({
            "status": "extracted",
            "message": f"Found {len(core_terms)} core terms. Please review and edit.",
            "terms": [{"english": term, "vietnamese": ""} for term in core_terms],
            "raw_text": raw_text,
            "filename": filename
        })
    finally:
        # Ensure temporary file cleanup
        if os.path.exists(temp_path):
            os.remove(temp_path)

@app.route('/format', methods=['POST'])
def format_glossary():
    data = request.json or {}
    terms = data.get("terms", [])
    filename = data.get("filename", "glossary.docx")
    
    if not terms:
        return jsonify({"error": "Vocabulary list is empty"}), 400
        
    base_name, _ = os.path.splitext(filename)
    out_filename = f"Glossary_Template_{base_name}.xlsx"
    
    try:
        # Generate directly in BytesIO memory stream (no local disk write)
        bio = io.BytesIO()
        generate_excel_glossary(terms, bio)
        bio.seek(0)
        
        return send_file(
            bio,
            as_attachment=True,
            download_name=out_filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        return jsonify({"error": f"Formatting error: {str(e)}"}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5001)
