from flask import Flask, render_template, request, jsonify, redirect, url_for, send_file
import os
import pandas as pd
from datetime import datetime
import json
import random
import time

# Create Flask app
app = Flask(__name__)

# Demo data for testing
DEMO_DATA = [
    {
        "USN": "1AT22CS001",
        "Name": "ADITYA KUMAR",
        "Semester": "4",
        "Total": "625/800",
        "Result": "PASS",
        "SGPA": "8.75"
    },
    {
        "USN": "1AT22CS002",
        "Name": "AKSHAY SHARMA",
        "Semester": "4",
        "Total": "710/800",
        "Result": "PASS",
        "SGPA": "9.25"
    },
    {
        "USN": "1AT22CS003",
        "Name": "ANANYA PATEL",
        "Semester": "4",
        "Total": "590/800",
        "Result": "PASS",
        "SGPA": "8.00"
    },
    {
        "USN": "1AT22CS004",
        "Name": "ANIRUDH REDDY",
        "Semester": "4",
        "Total": "450/800",
        "Result": "FAIL",
        "SGPA": "5.50"
    },
    {
        "USN": "1AT22CS005",
        "Name": "BHAVANA SINGH",
        "Semester": "4",
        "Total": "680/800",
        "Result": "PASS",
        "SGPA": "8.85"
    }
]

# Demo logs for display
DEMO_LOGS = [
    "Starting process for 5 USNs",
    "Processing 1AT22CS001...",
    "Accessing VTU Results website",
    "Successfully extracted details for 1AT22CS001",
    "Processing 1AT22CS002...",
    "Accessing VTU Results website",
    "Successfully extracted details for 1AT22CS002",
    "Processing 1AT22CS003...",
    "Accessing VTU Results website",
    "Successfully extracted details for 1AT22CS003",
    "Processing 1AT22CS004...",
    "Accessing VTU Results website",
    "Successfully extracted details for 1AT22CS004",
    "Processing 1AT22CS005...",
    "Accessing VTU Results website",
    "Successfully extracted details for 1AT22CS005",
    "All USNs processed successfully",
    "Generating Excel file..."
]

# Global variable to store the last excel filename
last_excel_filename = None

@app.route('/')
def index():
    """Render the main page."""
    return render_template('index.html', current_year=datetime.now().year)

@app.route('/demo')
def demo():
    """Render the demo page."""
    return render_template('demo.html', current_year=datetime.now().year)

@app.route('/api/check_api_key', methods=['GET'])
def check_api_key():
    """Check if 2Captcha API key is configured."""
    return jsonify({'api_key_configured': False})

@app.route('/api/run_script', methods=['POST'])
def run_script():
    """API endpoint for running the script."""
    try:
        data = request.json
        start_usn = data.get('start_usn')
        end_usn = data.get('end_usn')
        interactive_mode = data.get('interactive_mode', False)
        
        if not start_usn or not end_usn:
            return jsonify({'error': 'Start and end USN required'}), 400
        
        # This is a demo version, so we'll just use the demo data
        # In a real app, this would call the scraping functionality
        
        # Generate a unique filename for this session
        global last_excel_filename
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        last_excel_filename = f"results_{timestamp}.xlsx"
        
        # Simulate processing delay
        time.sleep(2)
        
        # Create DataFrame and save to Excel
        df = pd.DataFrame(DEMO_DATA)
        df.to_excel(last_excel_filename, index=False)
        
        # Return success response with demo data
        return jsonify({
            'status': 'success',
            'message': 'Results scraped successfully (Demo Mode)',
            'data': DEMO_DATA,
            'filename': last_excel_filename,
            'logs': DEMO_LOGS
        })
        
    except Exception as e:
        return jsonify({
            'status': 'error',
            'message': f'An unexpected error occurred: {str(e)}'
        }), 500

@app.route('/download/<filename>')
def download_file(filename):
    """Download the Excel file."""
    try:
        return send_file(filename, as_attachment=True)
    except Exception as e:
        return jsonify({
            'status': 'error',
            'message': f'Error downloading file: {str(e)}'
        }), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True) 