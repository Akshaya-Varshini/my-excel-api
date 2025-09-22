#!/usr/bin/env python3

"""
SIMPLE Flask API for Render Deployment - Fixed Version
"""

import os
from flask import Flask, request, jsonify
from flask_cors import CORS
import tempfile
import logging
from werkzeug.utils import secure_filename
import time

# Try to import your financial report generator
try:
    from financial_report_generator import EnhancedFinancialReportGenerator
    GENERATOR_AVAILABLE = True
except ImportError as e:
    print(f"Warning: Could not import financial_report_generator: {e}")
    GENERATOR_AVAILABLE = False

app = Flask(__name__)
CORS(app)

# Simple logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Configuration
UPLOAD_FOLDER = tempfile.gettempdir()
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
MAX_FILE_SIZE = 16 * 1024 * 1024  # 16MB

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = MAX_FILE_SIZE

# Get API keys from environment variables
GEMINI_KEY = os.environ.get('GEMINI_API_KEY')
PDFCO_KEY = os.environ.get('PDFCO_API_KEY')

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def home():
    """Simple home page"""
    return jsonify({
        'status': 'healthy',
        'message': 'Excel Processing API is running on Render!',
        'time': time.strftime('%Y-%m-%d %H:%M:%S'),
        'generator_available': GENERATOR_AVAILABLE,
        'api_keys_configured': {
            'gemini': bool(GEMINI_KEY),
            'pdfco': bool(PDFCO_KEY)
        },
        'endpoints': {
            'home': '/',
            'process': '/process'
        }
    })

@app.route('/process', methods=['POST'])
def process_files():
    """Process uploaded Excel files"""
    try:
        logger.info("Processing files...")

        # Check for required files
        required_files = ['balance_sheet', 'cash_flow', 'profit_loss']
        uploaded_files = {}

        for file_key in required_files:
            if file_key not in request.files:
                return jsonify({
                    'success': False,
                    'error': f'Missing file: {file_key}'
                }), 400

            file = request.files[file_key]
            if file.filename == '':
                return jsonify({
                    'success': False,
                    'error': f'No file selected for {file_key}'
                }), 400

            if not allowed_file(file.filename):
                return jsonify({
                    'success': False,
                    'error': f'Invalid file type for {file_key}. Use .xlsx or .xls'
                }), 400

            # Save file
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], f"{file_key}_{filename}")
            file.save(filepath)
            uploaded_files[file_key] = filepath

        # Check if we can use the real generator
        if not GENERATOR_AVAILABLE:
            # Clean up files
            for filepath in uploaded_files.values():
                try:
                    os.remove(filepath)
                except:
                    pass
            
            return jsonify({
                'success': True,
                'message': 'Files processed successfully (demo mode - generator not available)',
                'result': {
                    'status': 'success',
                    'message': 'Files uploaded successfully but generator module not available',
                    'timestamp': time.strftime('%Y-%m-%d %H:%M:%S'),
                    'files_received': list(uploaded_files.keys())
                },
                'timestamp': time.strftime('%Y-%m-%d %H:%M:%S')
            })

        # Check API keys
        if not GEMINI_KEY or not PDFCO_KEY:
            # Clean up files
            for filepath in uploaded_files.values():
                try:
                    os.remove(filepath)
                except:
                    pass
            
            return jsonify({
                'success': False,
                'error': 'API keys not configured. Please set GEMINI_API_KEY and PDFCO_API_KEY environment variables.'
            }), 500

        # Process with your generator - FIXED METHOD NAME
        logger.info("Initializing generator...")
        generator = EnhancedFinancialReportGenerator(GEMINI_KEY, PDFCO_KEY)
        
        logger.info("Processing comprehensive financial report...")
        # Use the correct method name from financial_report_generator.py
        html_content, pdf_url, chart_urls = generator.process_comprehensive_financial_report({
            'balance_sheet': uploaded_files['balance_sheet'],
            'cash_flow': uploaded_files['cash_flow'],
            'profit_loss': uploaded_files['profit_loss']
        })

        # Clean up files
        for filepath in uploaded_files.values():
            try:
                os.remove(filepath)
            except:
                pass

        return jsonify({
            'success': True,
            'message': 'Financial report generated successfully!',
            'result': {
                'status': 'success',
                'message': 'Financial report generated successfully',
                'html_content': html_content[:1000] + "..." if len(html_content) > 1000 else html_content,  # Truncate for response size
                'pdf_url': pdf_url,
                'chart_urls': chart_urls,
                'charts_generated': len([url for url in chart_urls if url]),
                'timestamp': time.strftime('%Y-%m-%d %H:%M:%S')
            },
            'timestamp': time.strftime('%Y-%m-%d %H:%M:%S')
        })

    except Exception as e:
        logger.error(f"Error: {str(e)}")
        
        # Clean up files on error
        if 'uploaded_files' in locals():
            for filepath in uploaded_files.values():
                try:
                    os.remove(filepath)
                except:
                    pass
        
        return jsonify({
            'success': False,
            'error': f'Processing failed: {str(e)}',
            'timestamp': time.strftime('%Y-%m-%d %H:%M:%S')
        }), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)