from flask import Flask, render_template, request, jsonify, send_file
import os
import tempfile
from werkzeug.utils import secure_filename
from document_parser import DocumentParser
from change_detector import ChangeDetector
from report_generator import ReportGenerator
import json

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['REPORTS_FOLDER'] = 'reports'

# Create necessary directoriess
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['REPORTS_FOLDER'], exist_ok=True)

ALLOWED_EXTENSIONS = {'docx', 'pdf', 'xlsx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/compare', methods=['POST'])
def compare_documents():
    try:
        if 'original' not in request.files or 'revised' not in request.files:
            return jsonify({'error': 'Both original and revised files are required'}), 400
        
        original_file = request.files['original']
        revised_file = request.files['revised']
        
        if original_file.filename == '' or revised_file.filename == '':
            return jsonify({'error': 'No files selected'}), 400
        
        if not (allowed_file(original_file.filename) and allowed_file(revised_file.filename)):
            return jsonify({'error': 'Invalid file type. Only DOCX, PDF, and XLSX files are allowed'}), 400
        
        # Save uploaded files
        original_filename = secure_filename(original_file.filename)
        revised_filename = secure_filename(revised_file.filename)
        
        original_path = os.path.join(app.config['UPLOAD_FOLDER'], original_filename)
        revised_path = os.path.join(app.config['UPLOAD_FOLDER'], revised_filename)
        
        original_file.save(original_path)
        revised_file.save(revised_path)
        
        # Parse documents
        parser = DocumentParser()
        original_content = parser.parse_document(original_path)
        revised_content = parser.parse_document(revised_path)
        
        # Detect changes
        detector = ChangeDetector()
        changes = detector.detect_changes(original_content, revised_content)
        
        # Generate report
        report_generator = ReportGenerator()
        report_path = report_generator.generate_report(changes, original_content, revised_content)
        
        # Clean up uploaded files
        os.remove(original_path)
        os.remove(revised_path)
        
        return jsonify({
            'success': True,
            'changes': changes,
            'report_path': report_path
        })
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/download_report/<filename>')
def download_report(filename):
    try:
        return send_file(os.path.join(app.config['REPORTS_FOLDER'], filename), as_attachment=True)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
