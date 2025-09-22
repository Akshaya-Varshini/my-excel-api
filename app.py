#!/usr/bin/env python3
"""
Flask API wrapper for Enhanced Financial Report Generator
"""

from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import os
import tempfile
import logging
from werkzeug.utils import secure_filename

# Import your existing financial report generator
from financial_report_generator import EnhancedFinancialReportGenerator

app = Flask(__name__)
CORS(app)  # Enable CORS for frontend communication

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Configuration
UPLOAD_FOLDER = tempfile.gettempdir()
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
MAX_FILE_SIZE = 16 * 1024 * 1024  # 16MB max file size

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = MAX_FILE_SIZE

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET'])
def health_check():
    """Health check endpoint"""
    return jsonify({
        'status': 'healthy',
        'message': 'Enhanced Financial Report Generator API is running'
    })

@app.route('/process', methods=['POST'])
def process_financial_files():
    """
    Process uploaded financial files and generate report
    """
    try:
        # Check if required files are present
        required_files = ['balance_sheet', 'cash_flow', 'profit_loss']
        uploaded_files = {}
        
        for file_type in required_files:
            if file_type not in request.files:
                return jsonify({
                    'error': f'Missing required file: {file_type}',
                    'success': False
                }), 400
                
            file = request.files[file_type]
            if file.filename == '':
                return jsonify({
                    'error': f'No file selected for: {file_type}',
                    'success': False
                }), 400
                
            if not allowed_file(file.filename):
                return jsonify({
                    'error': f'Invalid file type for {file_type}. Only .xlsx and .xls files are allowed.',
                    'success': False
                }), 400
                
            # Save uploaded file to temp directory
            filename = secure_filename(f"{file_type}_{file.filename}")
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            uploaded_files[file_type] = filepath
            
        logger.info(f"Files uploaded successfully: {list(uploaded_files.keys())}")
        
        # Initialize the financial report generator
        generator = EnhancedFinancialReportGenerator()
        
        # Process the files using your existing logic
        logger.info("Processing financial data...")
        html_content, pdf_url, chart_urls = generator.process_comprehensive_financial_report(uploaded_files)
        
        # Clean up uploaded files
        for filepath in uploaded_files.values():
            try:
                if os.path.exists(filepath):
                    os.remove(filepath)
            except Exception as e:
                logger.warning(f"Could not remove temp file {filepath}: {e}")
        
        # Return the generated report
        return jsonify({
            'success': True,
            'html_content': html_content,
            'pdf_url': pdf_url,
            'chart_urls': chart_urls,
            'message': 'Financial report generated successfully'
        })
        
    except Exception as e:
        logger.error(f"Error processing files: {e}")
        
        # Clean up any uploaded files on error
        if 'uploaded_files' in locals():
            for filepath in uploaded_files.values():
                try:
                    if os.path.exists(filepath):
                        os.remove(filepath)
                except:
                    pass
                    
        return jsonify({
            'error': str(e),
            'success': False,
            'message': 'Failed to process financial files'
        }), 500

@app.route('/download-pdf/<path:pdf_url>')
def download_pdf(pdf_url):
    """
    Proxy endpoint to download PDF (optional, for CORS handling)
    """
    try:
        import requests
        response = requests.get(pdf_url)
        response.raise_for_status()
        
        # Create temporary file for PDF
        temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        temp_pdf.write(response.content)
        temp_pdf.close()
        
        return send_file(
            temp_pdf.name,
            as_attachment=True,
            download_name='financial_report.pdf',
            mimetype='application/pdf'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.errorhandler(413)
def too_large(e):
    return jsonify({
        'error': 'File too large. Maximum size is 16MB per file.',
        'success': False
    }), 413

if __name__ == '__main__':
    print("🚀 Starting Enhanced Financial Report Generator API...")
    print("📊 Upload your financial files at: http://localhost:5000")
    print("🔗 API endpoint: http://localhost:5000/process")
    print("=" * 60)
    
    app.run(debug=True, host='0.0.0.0', port=5000)