from flask import Flask, request, jsonify, render_template, send_file
from flask_cors import CORS
import os
import sys
import pandas as pd
from datetime import datetime
import traceback
import subprocess

# Import original selenium_vtu_results script functions
import selenium_vtu_results

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes

@app.route('/')
def index():
    return render_template('index.html', current_year=datetime.now().year)

@app.route('/api/scrape', methods=['POST'])
def scrape():
    try:
        data = request.json
        start_usn = data.get('start_usn')
        end_usn = data.get('end_usn')
        
        if not start_usn or not end_usn:
            return jsonify({'error': 'Start and end USN required'}), 400
        
        # Parse USN numbers
        if start_usn.startswith("1AT22CS") and len(start_usn) == 10:
            start_num = int(start_usn[7:10])
        else:
            try:
                start_num = int(start_usn)
            except ValueError:
                return jsonify({'error': 'Invalid start USN format'}), 400
        
        if end_usn.startswith("1AT22CS") and len(end_usn) == 10:
            end_num = int(end_usn[7:10])
        else:
            try:
                end_num = int(end_usn)
            except ValueError:
                return jsonify({'error': 'Invalid end USN format'}), 400
        
        # Validate range
        if start_num > end_num:
            start_num, end_num = end_num, start_num
        
        # Generate USN list - without any limits
        usn_list = [f"1AT22CS{str(i).zfill(3)}" for i in range(start_num, end_num + 1)]
        
        print(f"Processing {len(usn_list)} USNs from {usn_list[0]} to {usn_list[-1]}")
        
        # Setup driver
        driver = selenium_vtu_results.setup_driver()
        
        try:
            # Process results using the original script
            all_results = selenium_vtu_results.process_results(driver, usn_list)
            
            # Save results to Excel
            excel_filename = selenium_vtu_results.save_to_excel(all_results)
            
            if excel_filename and os.path.exists(excel_filename):
                df = pd.read_excel(excel_filename)
                return jsonify({
                    'status': 'success',
                    'message': f'Scraped {len(df)} results',
                    'data': df.to_dict(orient='records'),
                    'filename': os.path.basename(excel_filename)
                })
            else:
                return jsonify({
                    'status': 'error',
                    'message': 'No results found or error occurred'
                }), 400
        finally:
            # Close the driver
            if driver:
                driver.quit()
            
    except Exception as e:
        return jsonify({
            'status': 'error',
            'message': str(e),
            'traceback': traceback.format_exc()
        }), 500

@app.route('/download/<filename>')
def download_file(filename):
    try:
        # Ensure filename doesn't contain path traversal
        if '..' in filename or '/' in filename:
            return jsonify({'error': 'Invalid filename'}), 400
            
        # Check if file exists
        if not os.path.exists(filename):
            return jsonify({'error': 'File not found'}), 404
            
        return send_file(
            filename,
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        return jsonify({
            'status': 'error',
            'message': str(e)
        }), 500

@app.route('/api/run_script', methods=['POST'])
def run_script():
    """Run the original script as a subprocess."""
    try:
        data = request.json
        start_usn = data.get('start_usn')
        end_usn = data.get('end_usn')
        
        if not start_usn or not end_usn:
            return jsonify({'error': 'Start and end USN required'}), 400
        
        # Run the script as a subprocess with the start and end USNs
        cmd = [sys.executable, 'selenium_vtu_results.py', start_usn, end_usn]
        process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        stdout, stderr = process.communicate()
        
        # Check for Excel files
        excel_files = [f for f in os.listdir('.') if f.startswith('vtu_results_') and f.endswith('.xlsx')]
        excel_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
        
        if excel_files and os.path.exists(excel_files[0]):
            df = pd.read_excel(excel_files[0])
            return jsonify({
                'status': 'success',
                'message': f'Scraped {len(df)} results',
                'data': df.to_dict(orient='records'),
                'filename': excel_files[0]
            })
        else:
            return jsonify({
                'status': 'error',
                'message': 'No results found or error occurred',
                'stdout': stdout,
                'stderr': stderr
            }), 400
            
    except Exception as e:
        return jsonify({
            'status': 'error',
            'message': str(e),
            'traceback': traceback.format_exc()
        }), 500

@app.route('/jump_to_usn', methods=['POST'])
def jump_to_usn():
    try:
        data = request.json
        target_usn = data.get('targetUsn')
        
        if not target_usn:
            return jsonify({
                'success': False,
                'message': 'No USN provided'
            }), 400
        
        # Use vtu_results_hosted.py's jump_to_usn function if it's available
        try:
            from vtu_results_hosted import processing_active, processing_function
            
            if not processing_active or not processing_function:
                # If not active, we'll proceed with our standalone implementation
                raise ImportError("No active processing session")
                
            result = processing_function['jump_to_usn'](target_usn)
            
            if result:
                return jsonify({
                    'success': True,
                    'message': f'Successfully jumped to USN: {target_usn}'
                })
            else:
                return jsonify({
                    'success': False,
                    'message': f'Failed to jump to USN: {target_usn} - not found in range or invalid format'
                })
                
        except (ImportError, NameError):
            # Standalone implementation
            # Format USN if needed
            if not target_usn.startswith("1AT22CS") and target_usn.isdigit():
                try:
                    usn_num = int(target_usn)
                    target_usn = f"1AT22CS{str(usn_num).zfill(3)}"
                except ValueError:
                    return jsonify({
                        'success': False,
                        'message': 'Invalid USN format'
                    }), 400
            
            # Setup driver
            driver = selenium_vtu_results.setup_driver()
            
            try:
                # Process just this single USN
                results = selenium_vtu_results.process_results(driver, [target_usn])
                
                if results and len(results) > 0:
                    # Check if results is a list of dictionaries (correct structure)
                    if isinstance(results, list):
                        # Save to Excel
                        excel_filename = selenium_vtu_results.save_to_excel(results)
                        
                        return jsonify({
                            'success': True,
                            'message': f'Successfully found and jumped to USN: {target_usn}',
                            'data': results[0],  # First item for display
                            'filename': os.path.basename(excel_filename) if excel_filename else None
                        })
                    else:
                        # Handle unexpected result format
                        return jsonify({
                            'success': False,
                            'message': f'Invalid results format for USN: {target_usn}'
                        })
            finally:
                # Close the driver
                if driver:
                    driver.quit()
                    
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'Error processing USN: {str(e)}'
        }), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port) 