import time
import pandas as pd
from datetime import datetime
import os
import base64
import requests
import json
import traceback
import sys
import re
from flask import Flask, request, jsonify, render_template, send_file, redirect, url_for
from flask_cors import CORS
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# 2Captcha API key - Replace with your actual API key
API_KEY = "20480f95adb6216bc0e788f58c343c11"  # 2Captcha API key

# Initialize Flask app
app = Flask(__name__)
CORS(app)  # Enable CORS for all routes

# Global variables to control processing flow
SKIP_CURRENT_USN = False
EXIT_PROCESSING = False
LAST_EXCEL_FILENAME = None  # Track the last generated Excel file

def setup_driver():
    """Set up and return a Chrome WebDriver instance."""
    try:
        chrome_options = Options()
        chrome_options.add_argument("--start-maximized")
        
        # For server environments, add headless options (unless in manual CAPTCHA mode)
        if not os.environ.get('DEVELOPMENT') and not os.environ.get('MANUAL_CAPTCHA'):
            chrome_options.add_argument("--headless=new")
            chrome_options.add_argument("--disable-gpu")
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--remote-debugging-port=9222")
        else:
            # Always add these options for more reliable operation in cloud environments
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
        
        # Add window size to ensure visibility of elements
        chrome_options.add_argument("--window-size=1920,1080")
        
        # Check if we're running on Render.com
        if os.environ.get('RENDER'):
            print("Running on Render.com, using special configuration")
            chrome_options.add_argument("--disable-extensions")
            chrome_options.add_argument("--disable-setuid-sandbox")
            chrome_options.binary_location = "/opt/google/chrome/chrome"
        
        # Detect operating system
        if sys.platform.startswith('win'):
            # Windows - use ChromeDriverManager
            try:
                service = Service(ChromeDriverManager().install())
                driver = webdriver.Chrome(service=service, options=chrome_options)
            except Exception as e:
                print(f"Error with ChromeDriverManager: {str(e)}")
                # Fallback to direct Chrome initialization
                driver = webdriver.Chrome(options=chrome_options)
        else:
            # Linux/Mac - likely running in a container
            try:
                service = Service(ChromeDriverManager().install())
                driver = webdriver.Chrome(service=service, options=chrome_options)
            except Exception as e:
                print(f"Error with ChromeDriverManager: {str(e)}")
                print("Trying direct Chrome initialization")
                driver = webdriver.Chrome(options=chrome_options)
            
        print("Chrome WebDriver initialized successfully")
        return driver
    except Exception as e:
        print(f"Failed to initialize Chrome WebDriver: {str(e)}")
        traceback.print_exc()
        return None

def solve_captcha(driver):
    """Solve CAPTCHA using 2Captcha service."""
    if not API_KEY or API_KEY == "YOUR_2CAPTCHA_API_KEY":
        print("WARNING: No valid 2Captcha API key provided. You need to sign up at 2captcha.com and get an API key.")
        return None
        
    try:
        # Find the CAPTCHA image
        captcha_img = driver.find_element(By.CSS_SELECTOR, "img[alt='CAPTCHA code']")
        
        # Get the image source
        img_src = captcha_img.get_attribute("src")
        
        if "base64" in img_src:
            # Extract base64 encoded image
            img_base64 = img_src.split(",")[1]
        else:
            # Download the image
            response = requests.get(img_src, verify=False)
            img_base64 = base64.b64encode(response.content).decode("utf-8")
        
        # Send the CAPTCHA to 2Captcha
        data = {
            "key": API_KEY,
            "method": "base64",
            "body": img_base64,
            "json": 1
        }
        
        print("Sending CAPTCHA to 2Captcha service...")
        try:
            response = requests.post("https://2captcha.com/in.php", data=data)
            response_json = response.json()
            
            if response_json["status"] == 1:
                request_id = response_json["request"]
                print(f"CAPTCHA sent successfully. Request ID: {request_id}")
                
                # Wait for the CAPTCHA to be solved
                print("Waiting for 2Captcha to solve the CAPTCHA...")
                for attempt in range(30):  # Try for 30 attempts (about 150 seconds)
                    time.sleep(5)
                    try:
                        get_url = f"https://2captcha.com/res.php?key={API_KEY}&action=get&id={request_id}&json=1"
                        response = requests.get(get_url)
                        response_json = response.json()
                        
                        if response_json["status"] == 1:
                            captcha_text = response_json["request"]
                            print(f"CAPTCHA solved: {captcha_text}")
                            return captcha_text
                        elif "CAPCHA_NOT_READY" in response_json["request"]:
                            print(f"CAPTCHA not ready yet. Attempt {attempt+1}/30...")
                        else:
                            print(f"Error getting CAPTCHA solution: {response_json['request']}")
                            return None
                    except Exception as e:
                        print(f"Error checking CAPTCHA solution: {str(e)}")
                
                print("Timeout waiting for CAPTCHA solution")
                return None
            else:
                error_msg = response_json["request"]
                print(f"Error sending CAPTCHA to 2Captcha: {error_msg}")
                
                if "ERROR_KEY_DOES_NOT_EXIST" in error_msg:
                    print("The API key does not exist or is invalid. Please check your 2Captcha API key.")
                elif "ERROR_ZERO_BALANCE" in error_msg:
                    print("Your 2Captcha account has no balance. Please add funds to your account.")
                elif "ERROR_NO_SLOT_AVAILABLE" in error_msg:
                    print("No slots available on 2Captcha servers. Try again later.")
                elif "ERROR_ZERO_CAPTCHA_FILESIZE" in error_msg:
                    print("The CAPTCHA image could not be loaded or is empty.")
                else:
                    print("Unknown error from 2Captcha service. Check their documentation for more details.")
                
                return None
        except Exception as e:
            print(f"Error communicating with 2Captcha service: {str(e)}")
            return None
    except Exception as e:
        print(f"Error preparing CAPTCHA for solving: {str(e)}")
        return None 

def process_results(driver, usn_list, manual_mode=False):
    """Process results for a list of USNs."""
    global SKIP_CURRENT_USN, EXIT_PROCESSING
    
    # Reset control flags at the start of processing
    SKIP_CURRENT_USN = False
    EXIT_PROCESSING = False
    
    base_url = "https://results.vtu.ac.in/DJcbcs25/index.php"
    all_results = []
    processing_logs = []
    
    i = 0
    max_retries = 3  # Maximum number of CAPTCHA retries for each USN
    retries = 0      # Current retry count
    
    while i < len(usn_list):
        # Check if we should exit processing early
        if EXIT_PROCESSING:
            log_message = "Process terminated by user. Saving partial results."
            print(log_message)
            processing_logs.append(log_message)
            break
            
        # Reset the skip flag for the new USN
        if retries == 0:  # Only reset if we're not in retry mode
            SKIP_CURRENT_USN = False
        
        usn = usn_list[i]
        
        # Show retry information if applicable
        if retries > 0:
            log_message = f"\nRetrying {usn} (Attempt {retries+1}/{max_retries+1})..."
        else:
            log_message = f"\nProcessing {usn} ({i+1}/{len(usn_list)})..."
            
        print(log_message)
        processing_logs.append(log_message)
        
        # Navigate to the results page
        try:
            driver.get(base_url)
            time.sleep(2)  # Wait for page to load
        except Exception as e:
            log_message = f"Error loading page: {str(e)}"
            print(log_message)
            processing_logs.append(log_message)
            i += 1  # Move to next USN
            retries = 0  # Reset retry counter
            continue
        
        # Try different selectors for the USN input field
        usn_input = None
        selectors = [
            "input[placeholder='ENTER USN']",
            "input[name='lns']",
            "input.form-control[type='text']",
            "input[minlength='10'][maxlength='10']"
        ]
        
        for selector in selectors:
            try:
                usn_input = driver.find_element(By.CSS_SELECTOR, selector)
                log_message = f"Found USN input field using selector: {selector}"
                print(log_message)
                processing_logs.append(log_message)
                break
            except NoSuchElementException:
                continue
        
        if not usn_input:
            log_message = f"Could not find USN input field for {usn} using any selector"
            print(log_message)
            processing_logs.append(log_message)
            i += 1  # Move to next USN
            retries = 0  # Reset retry counter
            continue
        
        # Enter USN
        usn_input.clear()
        usn_input.send_keys(usn)
        log_message = f"Entered USN: {usn}"
        print(log_message)
        processing_logs.append(log_message)
        
        # Check if we should skip this USN
        if SKIP_CURRENT_USN:
            log_message = f"User requested to skip USN: {usn}"
            print(log_message)
            processing_logs.append(log_message)
            i += 1  # Move to next USN
            retries = 0  # Reset retry counter
            continue
        
        # Find CAPTCHA input field
        captcha_input = None
        try:
            captcha_input = driver.find_element(By.CSS_SELECTOR, "input[name='captchacode']")
            log_message = "Found CAPTCHA input field"
            print(log_message)
            processing_logs.append(log_message)
        except NoSuchElementException:
            log_message = "Could not find CAPTCHA input field"
            print(log_message)
            processing_logs.append(log_message)
            i += 1  # Move to next USN
            retries = 0  # Reset retry counter
            continue
        
        captcha_text = None
        
        # In manual mode, we let the user type the captcha
        if manual_mode:
            log_message = "Manual mode enabled. Please look at the browser window and type the CAPTCHA and click Submit."
            print(log_message)
            processing_logs.append(log_message)
            
            # Make the browser window visible if it's in headless mode
            if not os.environ.get('DEVELOPMENT'):
                # Use JavaScript to make the window visible
                driver.execute_script("document.body.style.opacity = '1';")
            
            # Prompt the user to enter the CAPTCHA
            print("Please enter the CAPTCHA shown in the browser window and click Submit:")
            
            # Give focus to the CAPTCHA input field in the browser
            driver.execute_script("arguments[0].focus();", captcha_input)
            
            # Wait for the page to change (user submits the form)
            current_url = driver.current_url
            page_changed = False
            invalid_captcha = False
            captcha_timeout = time.time() + 60
            
            while time.time() < captcha_timeout and not page_changed and not invalid_captcha:
                try:
                    # Check for alert (invalid captcha)
                    try:
                        alert = driver.switch_to.alert
                        alert_text = alert.text
                        log_message = f"Alert detected: {alert_text}"
                        print(log_message)
                        processing_logs.append(log_message)
                        
                        if "Invalid captcha" in alert_text:
                            log_message = f"Invalid CAPTCHA detected for {usn}. Moving to next USN."
                            print(log_message)
                            processing_logs.append(log_message)
                            alert.accept()  # Dismiss the alert
                            invalid_captcha = True
                    except:
                        # No alert present, continue checking
                        pass
                    
                    # Check if the page has changed or if we're no longer on the CAPTCHA page
                    if current_url != driver.current_url or "captchacode" not in driver.page_source:
                        page_changed = True
                        log_message = "Page changed, user submitted the form successfully."
                        print(log_message)
                        processing_logs.append(log_message)
                        break
                    time.sleep(0.5)  # Check every half second
                except Exception as e:
                    log_message = f"Error checking page state: {str(e)}"
                    print(log_message)
                    processing_logs.append(log_message)
                    # Don't break here, just log the error and continue
                    time.sleep(0.5)
            
            # If invalid captcha was detected, move to next USN
            if invalid_captcha:
                i += 1  # Move to next USN
                retries = 0  # Reset retry counter
                continue
                
            # If timeout was reached
            if not page_changed:
                log_message = "Input timeout. Moving to next USN."
                print(log_message)
                processing_logs.append(log_message)
                i += 1  # Move to next USN
                retries = 0  # Reset retry counter
                continue
            
            # If we got here, the user has submitted the form, we don't need to click the submit button
            # Skip to waiting for results page to load
            log_message = "Waiting for results page to load..."
            print(log_message)
            processing_logs.append(log_message)
            time.sleep(3)
            
            # Check if we're still on the input page (CAPTCHA error)
            if "ENTER USN" in driver.page_source or "captchacode" in driver.page_source:
                if retries < max_retries:
                    retries += 1
                    log_message = f"CAPTCHA validation failed for {usn}. Will retry (Attempt {retries+1}/{max_retries+1})."
                    print(log_message)
                    processing_logs.append(log_message)
                    continue  # Retry the same USN
                else:
                    log_message = f"Maximum CAPTCHA retries ({max_retries+1}) reached for {usn}. Moving to next USN."
                    print(log_message)
                    processing_logs.append(log_message)
                    i += 1  # Move to next USN
                    retries = 0  # Reset retry counter
                    continue
            
            # Reset retry counter as we've passed CAPTCHA validation
            retries = 0
            
            # Continue to extract results
        else:
            # Try to solve CAPTCHA using 2Captcha
            captcha_text = solve_captcha(driver)
            
            if captcha_text:
                log_message = f"Automatically solved CAPTCHA: {captcha_text}"
                print(log_message)
                processing_logs.append(log_message)
                
                # Enter the CAPTCHA text
                captcha_input.clear()
                captcha_input.send_keys(captcha_text)
            else:
                log_message = "Automatic CAPTCHA solving failed. Skipping this USN."
                print(log_message)
                processing_logs.append(log_message)
                i += 1  # Move to next USN
                retries = 0  # Reset retry counter
                continue
            
            # Find and click the submit button
            try:
                submit_button = driver.find_element(By.CSS_SELECTOR, "input[type='submit']")
                submit_button.click()
                log_message = "Clicked submit button"
                print(log_message)
                processing_logs.append(log_message)
            except NoSuchElementException:
                log_message = "Could not find submit button"
                print(log_message)
                processing_logs.append(log_message)
                i += 1  # Move to next USN
                retries = 0  # Reset retry counter
                continue
            
            # Wait for results page to load
            log_message = "Waiting for results page to load..."
            print(log_message)
            processing_logs.append(log_message)
            time.sleep(3)
            
            # Check if we're still on the input page (CAPTCHA error)
            if "ENTER USN" in driver.page_source or "captchacode" in driver.page_source:
                if retries < max_retries:
                    retries += 1
                    log_message = f"CAPTCHA validation failed for {usn}. Will retry (Attempt {retries+1}/{max_retries+1})."
                    print(log_message)
                    processing_logs.append(log_message)
                    continue  # Retry the same USN
                else:
                    log_message = f"Maximum CAPTCHA retries ({max_retries+1}) reached for {usn}. Moving to next USN."
                    print(log_message)
                    processing_logs.append(log_message)
                    i += 1  # Move to next USN
                    retries = 0  # Reset retry counter
                    continue
            
            # Reset retry counter as we've passed CAPTCHA validation
            retries = 0
        
        # Extract results
        try:
            # Extract student info
            student_info = {}
            student_info["USN"] = usn  # Default to input USN
            
            # Get USN and student name from the table
            try:
                # Look for the table with student information
                tables = driver.find_elements(By.TAG_NAME, "table")
                for table in tables:
                    try:
                        rows = table.find_elements(By.TAG_NAME, "tr")
                        for row in rows:
                            cells = row.find_elements(By.TAG_NAME, "td")
                            if len(cells) >= 2:
                                cell_text = cells[0].text.strip()
                                if "University Seat Number" in cell_text:
                                    usn_text = cells[1].text.strip()
                                    if ":" in usn_text:
                                        usn_text = usn_text.split(":", 1)[1].strip()
                                    student_info["USN"] = usn_text
                                    log_message = f"Found USN: {usn_text}"
                                    print(log_message)
                                    processing_logs.append(log_message)
                                elif "Student Name" in cell_text:
                                    name_text = cells[1].text.strip()
                                    if ":" in name_text:
                                        name_text = name_text.split(":", 1)[1].strip()
                                    student_info["Student Name"] = name_text
                                    log_message = f"Found Student Name: {name_text}"
                                    print(log_message)
                                    processing_logs.append(log_message)
                    except Exception as e:
                        log_message = f"Error processing table row: {str(e)}"
                        print(log_message)
                        processing_logs.append(log_message)
            except Exception as e:
                log_message = f"Error finding student info table: {str(e)}"
                print(log_message)
                processing_logs.append(log_message)
            
            # Get semester
            try:
                # Look for the div with semester information
                semester_divs = driver.find_elements(By.XPATH, "//div[contains(text(), 'Semester')]")
                if not semester_divs:
                    # Try another approach with CSS selector
                    semester_divs = driver.find_elements(By.CSS_SELECTOR, "div[style*='text-align:center']")
                
                for div in semester_divs:
                    try:
                        semester_text = div.text.strip()
                        # Extract the semester number
                        import re
                        semester_match = re.search(r'Semester\s*:\s*(\d+)', semester_text)
                        if semester_match:
                            student_info["Semester"] = semester_match.group(1)
                            log_message = f"Found Semester: {student_info['Semester']}"
                            print(log_message)
                            processing_logs.append(log_message)
                            break
                    except Exception as e:
                        log_message = f"Error processing semester div: {str(e)}"
                        print(log_message)
                        processing_logs.append(log_message)
                
                # If semester is still not found, try one more approach
                if "Semester" not in student_info:
                    # Try to find it in the page source
                    page_source = driver.page_source
                    semester_match = re.search(r'Semester\s*:\s*(\d+)', page_source)
                    if semester_match:
                        student_info["Semester"] = semester_match.group(1)
                        log_message = f"Found Semester from page source: {student_info['Semester']}"
                        print(log_message)
                        processing_logs.append(log_message)
            except Exception as e:
                log_message = f"Error finding semester: {str(e)}"
                print(log_message)
                processing_logs.append(log_message)
            
            # Initialize results dictionary with student info
            results = {}
            for key, value in student_info.items():
                results[key] = value
            
            # Find the divTable structure that contains the results
            try:
                # First, check if we can find the divTable directly
                div_tables = driver.find_elements(By.CLASS_NAME, "divTable")
                
                if div_tables:
                    log_message = f"Found {len(div_tables)} divTable elements"
                    print(log_message)
                    processing_logs.append(log_message)
                    
                    for div_table in div_tables:
                        # Check if this is the results table by looking for headers
                        try:
                            header_row = div_table.find_element(By.XPATH, ".//div[contains(@class, 'divTableRow')][1]")
                            header_cells = header_row.find_elements(By.CLASS_NAME, "divTableCell")
                            
                            header_texts = [cell.text.strip() for cell in header_cells]
                            log_message = f"Found table with headers: {header_texts}"
                            print(log_message)
                            processing_logs.append(log_message)
                            
                            if "Subject Code" in header_texts and "Subject Name" in header_texts:
                                log_message = "This is the results table!"
                                print(log_message)
                                processing_logs.append(log_message)
                                
                                # Get all rows except the header
                                data_rows = div_table.find_elements(By.XPATH, ".//div[contains(@class, 'divTableRow')][position() > 1]")
                                log_message = f"Found {len(data_rows)} subject rows"
                                print(log_message)
                                processing_logs.append(log_message)
                                
                                for row in data_rows:
                                    try:
                                        cells = row.find_elements(By.CLASS_NAME, "divTableCell")
                                        
                                        if len(cells) >= 6:  # Ensure we have enough cells
                                            subject_code = cells[0].text.strip()
                                            subject_name = cells[1].text.strip()
                                            internal_marks = cells[2].text.strip()
                                            external_marks = cells[3].text.strip()
                                            total_marks = cells[4].text.strip()
                                            result = cells[5].text.strip()
                                            
                                            log_message = f"Found subject: {subject_code} - {subject_name} - Internal: {internal_marks}, External: {external_marks}, Total: {total_marks}, Result: {result}"
                                            print(log_message)
                                            processing_logs.append(log_message)
                                            
                                            # Store in results dictionary
                                            results[subject_code] = {
                                                'Subject Name': subject_name,
                                                'Internal': internal_marks,
                                                'External': external_marks,
                                                'Total': total_marks,
                                                'Result': result
                                            }
                                    except Exception as e:
                                        log_message = f"Error processing row: {str(e)}"
                                        print(log_message)
                                        processing_logs.append(log_message)
                                
                                # We found and processed the results table, so break the loop
                                break
                        except Exception as e:
                            log_message = f"Error processing table: {str(e)}"
                            print(log_message)
                            processing_logs.append(log_message)
                
                # If we couldn't find results using divTable, try using XPath directly
                if len(results) <= 1:  # Only has USN, no subjects
                    log_message = "Could not find any subject results in the divTable structure"
                    print(log_message)
                    processing_logs.append(log_message)
                    log_message = "Trying alternative approach using XPath directly..."
                    print(log_message)
                    processing_logs.append(log_message)
                    
                    # Try to find all divTableRow elements that might contain subject data
                    subject_rows = driver.find_elements(By.XPATH, "//div[contains(@class, 'divTableRow')]")
                    
                    for row in subject_rows:
                        try:
                            cells = row.find_elements(By.CLASS_NAME, "divTableCell")
                            
                            if len(cells) >= 6:  # Ensure we have enough cells
                                # Check if first cell looks like a subject code (typically alphanumeric)
                                first_cell_text = cells[0].text.strip()
                                if first_cell_text and any(c.isalpha() for c in first_cell_text) and any(c.isdigit() for c in first_cell_text):
                                    subject_code = first_cell_text
                                    subject_name = cells[1].text.strip()
                                    internal_marks = cells[2].text.strip()
                                    external_marks = cells[3].text.strip()
                                    total_marks = cells[4].text.strip()
                                    result = cells[5].text.strip()
                                    
                                    log_message = f"Found subject (alt method): {subject_code} - {subject_name} - Internal: {internal_marks}, External: {external_marks}, Total: {total_marks}, Result: {result}"
                                    print(log_message)
                                    processing_logs.append(log_message)
                                    
                                    # Store in results dictionary
                                    results[subject_code] = {
                                        'Subject Name': subject_name,
                                        'Internal': internal_marks,
                                        'External': external_marks,
                                        'Total': total_marks,
                                        'Result': result
                                    }
                        except Exception as e:
                            log_message = f"Error processing row (alt method): {str(e)}"
                            print(log_message)
                            processing_logs.append(log_message)
            
            except Exception as e:
                log_message = f"Error finding or processing divTable: {str(e)}"
                print(log_message)
                processing_logs.append(log_message)
            
            if len(results) > 1:  # At least USN and some other data
                all_results.append(results)
                log_message = f"Successfully extracted results for {usn} with {len(results)-1} subjects"
                print(log_message)
                processing_logs.append(log_message)
            else:
                log_message = f"No results data found for {usn}"
                print(log_message)
                processing_logs.append(log_message)
                
            # If we successfully extracted results, then increment i
            i += 1
        except Exception as e:
            log_message = f"Error processing results for {usn}: {str(e)}"
            print(log_message)
            processing_logs.append(log_message)
            i += 1  # Move to next USN
            retries = 0  # Reset retry counter
    
    return all_results, processing_logs 

def save_to_excel(all_results):
    """Save results to Excel file and return the filename."""
    global LAST_EXCEL_FILENAME
    
    if not all_results:
        print("\nNo results were collected!")
        return None
    
    print("\n" + "="*50)
    print(f"Processing {len(all_results)} results for Excel export...")
    print("="*50)
        
    # Create a normalized DataFrame
    # First, extract all possible subject codes
    all_subjects = set()
    for result in all_results:
        for key in result.keys():
            if key not in ['USN', 'Student Name', 'Semester']:  # Subject codes
                all_subjects.add(key)
    
    print(f"Found {len(all_subjects)} unique subjects across all results")
    
    # Create rows for the DataFrame
    rows = []
    for result in all_results:
        row = {'USN': result.get('USN', '')}
        
        # Add student info
        for key, value in result.items():
            if key in ['Student Name', 'Semester']:
                row[key] = value
        
        # Add subject marks
        for subject in all_subjects:
            if subject in result:
                subject_data = result[subject]
                if isinstance(subject_data, dict):
                    row[f"{subject}_Name"] = subject_data.get('Subject Name', '')
                    row[f"{subject}_Internal"] = subject_data.get('Internal', '')
                    row[f"{subject}_External"] = subject_data.get('External', '')
                    row[f"{subject}_Total"] = subject_data.get('Total', '')
                    row[f"{subject}_Result"] = subject_data.get('Result', '')
        
        rows.append(row)
    
    # Create DataFrame
    df = pd.DataFrame(rows)
    
    # Save to Excel with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_filename = f"vtu_results_{timestamp}.xlsx"
    
    print(f"\nSaving data to Excel file: {excel_filename}")
    print(f"DataFrame shape: {df.shape[0]} rows × {df.shape[1]} columns")
    
    # Display a sample of the data
    if not df.empty:
        print("\nSample data (first few rows):")
        pd.set_option('display.max_columns', 10)
        pd.set_option('display.width', 1000)
        print(df.head(3).to_string())
    
    df.to_excel(excel_filename, index=False)
    print(f"\nResults successfully saved to {excel_filename}")
    
    # Save the filename globally
    LAST_EXCEL_FILENAME = excel_filename
    
    return excel_filename 

# Define Flask routes
@app.route('/')
def index():
    """Render the main page."""
    # Check if we should go directly to demo mode (e.g., if Selenium is not available)
    if os.environ.get('FORCE_DEMO') == 'True':
        return redirect(url_for('demo'))
    return render_template('index.html', current_year=datetime.now().year)

@app.route('/api/scrape', methods=['POST'])
def scrape():
    """API endpoint for scraping VTU results."""
    try:
        data = request.json
        start_usn = data.get('start_usn')
        end_usn = data.get('end_usn')
        interactive_mode = data.get('interactive_mode', False)
        
        if not start_usn or not end_usn:
            return jsonify({'error': 'Start and end USN required'}), 400
        
        # Parse USN numbers with flexible pattern
        usn_pattern = re.compile(r'^1AT(\d{2})(CS|IS|ME|EE|EC|CV|AI|ML)(\d{3})$')
        
        # Try to match the pattern for start USN
        start_match = usn_pattern.match(start_usn)
        if start_match:
            start_year = start_match.group(1)
            start_branch = start_match.group(2)
            start_num = int(start_match.group(3))
        else:
            try:
                # If no pattern match, try to extract just the number
                start_num = int(start_usn)
                start_year = "22"  # Default to current year
                start_branch = "CS"  # Default to CS branch
            except ValueError:
                return jsonify({'error': 'Invalid start USN format'}), 400
        
        # Try to match the pattern for end USN
        end_match = usn_pattern.match(end_usn)
        if end_match:
            end_year = end_match.group(1)
            end_branch = end_match.group(2)
            end_num = int(end_match.group(3))
            
            # Validate that years and branches match
            if end_year != start_year or end_branch != start_branch:
                return jsonify({'error': 'Start and end USN must have the same year and branch'}), 400
        else:
            try:
                # If no pattern match, try to extract just the number
                end_num = int(end_usn)
                if end_year != start_year or end_branch != start_branch:
                    return jsonify({'error': 'Start and end USN must have the same year and branch'}), 400
            except ValueError:
                return jsonify({'error': 'Invalid end USN format'}), 400
        
        # Validate range
        if start_num > end_num:
            start_num, end_num = end_num, start_num
        
        # Limit range for server-side processing to avoid timeouts
        if end_num - start_num > 10:
            end_num = start_num + 10
        
        # Generate USN list with the correct pattern
        usn_list = [f"1AT{start_year}{start_branch}{str(i).zfill(3)}" for i in range(start_num, end_num + 1)]
        
        # Check if we should use manual mode (no automatic CAPTCHA solving)
        manual_mode = interactive_mode or not API_KEY or API_KEY == "YOUR_2CAPTCHA_API_KEY"
        
        # If using manual mode, ensure we're in non-headless mode
        if manual_mode:
            os.environ['MANUAL_CAPTCHA'] = 'True'
        
        # Setup driver
        driver = setup_driver()
        if driver is None:
            return jsonify({
                'status': 'error',
                'message': 'Failed to initialize Chrome WebDriver. Please check if Chrome is installed correctly.'
            }), 500
        
        try:
            # Process results
            all_results, logs = process_results(driver, usn_list, manual_mode)
            
            if all_results:
                # Save results to Excel
                excel_filename = save_to_excel(all_results)
                
                if excel_filename and os.path.exists(excel_filename):
                    # Create a simplified version of results for the response
                    simplified_results = []
                    for result in all_results:
                        student = {
                            'USN': result.get('USN', ''),
                            'Name': result.get('Student Name', ''),
                            'Semester': result.get('Semester', ''),
                            'Subjects': []
                        }
                        
                        # Add subjects
                        total_marks = 0
                        for key, value in result.items():
                            if key not in ['USN', 'Student Name', 'Semester']:
                                if isinstance(value, dict):
                                    subject = {
                                        'Code': key,
                                        'Name': value.get('Subject Name', ''),
                                        'Internal': value.get('Internal', ''),
                                        'External': value.get('External', ''),
                                        'Total': value.get('Total', ''),
                                        'Result': value.get('Result', '')
                                    }
                                    student['Subjects'].append(subject)
                                    
                                    # Calculate total marks
                                    try:
                                        total_marks += int(value.get('Total', '0'))
                                    except ValueError:
                                        pass
                        
                        student['Total'] = str(total_marks)
                        student['Result'] = 'PASS' if all(s.get('Result', '') == 'P' for s in student['Subjects']) else 'FAIL'
                        simplified_results.append(student)
                    
                    return jsonify({
                        'status': 'success',
                        'message': f'Successfully scraped {len(simplified_results)} results',
                        'data': simplified_results,
                        'logs': logs,
                        'filename': excel_filename
                    })
                else:
                    return jsonify({
                        'status': 'error',
                        'message': 'Failed to save results to Excel',
                        'logs': logs
                    }), 500
            else:
                return jsonify({
                    'status': 'error',
                    'message': 'No results found',
                    'logs': logs
                }), 400
        finally:
            # Close the driver
            try:
                if driver:
                    driver.quit()
                    print("Driver closed successfully")
            except Exception as e:
                print(f"Error closing driver: {str(e)}")
            
    except Exception as e:
        error_details = traceback.format_exc()
        print(f"Error in scrape endpoint: {str(e)}")
        print(error_details)
        return jsonify({
            'status': 'error',
            'message': f"An unexpected error occurred: {str(e)}",
            'traceback': error_details
        }), 500

@app.route('/download/<filename>')
def download_file(filename):
    """Download the Excel file."""
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

@app.route('/api/check_api_key', methods=['GET'])
def check_api_key():
    """Check if a valid 2Captcha API key is configured."""
    if not API_KEY or API_KEY == "YOUR_2CAPTCHA_API_KEY" or API_KEY == "20480f95adb6216bc0e788f58c343c11":
        return jsonify({
            'status': 'warning',
            'message': 'No valid 2Captcha API key configured. CAPTCHA solving will require manual input.',
            'api_key_configured': False
        })
    else:
        return jsonify({
            'status': 'success', 
            'message': '2Captcha API key is configured. Automatic CAPTCHA solving is available.',
            'api_key_configured': True
        })

@app.route('/api/skip_usn', methods=['POST'])
def skip_usn():
    """Skip the currently processing USN."""
    global SKIP_CURRENT_USN
    SKIP_CURRENT_USN = True
    return jsonify({
        'status': 'success', 
        'message': 'Current USN will be skipped'
    })

@app.route('/api/exit_process', methods=['POST'])
def exit_process():
    """Exit the current processing and save partial results."""
    global EXIT_PROCESSING, LAST_EXCEL_FILENAME
    EXIT_PROCESSING = True
    
    response_data = {
        'status': 'success', 
        'message': 'Processing will be terminated and partial results saved'
    }
    
    # Include the filename if we have one
    if LAST_EXCEL_FILENAME and os.path.exists(LAST_EXCEL_FILENAME):
        response_data['filename'] = LAST_EXCEL_FILENAME
    
    return jsonify(response_data)

@app.route('/demo', methods=['GET'])
def demo():
    """Demo page that shows sample results without using Selenium."""
    return render_template('demo.html', current_year=datetime.now().year)

@app.route('/api/demo', methods=['POST'])
def demo_data():
    """Return sample results data for demo purposes."""
    try:
        data = request.json
        start_usn = data.get('start_usn', '1')
        end_usn = data.get('end_usn', '5')
        
        # Create sample results
        sample_results = []
        for i in range(int(start_usn), int(end_usn) + 1):
            usn = f"1AT22CS{str(i).zfill(3)}"
            student = {
                'USN': usn,
                'Name': f"Demo Student {i}",
                'Semester': "5",
                'Total': str(350 + i * 10),
                'Result': 'PASS' if i % 5 != 0 else 'FAIL',
                'Subjects': [
                    {
                        'Code': '18CS51',
                        'Name': 'Management and Entrepreneurship',
                        'Internal': '18',
                        'External': '52',
                        'Total': '70',
                        'Result': 'P'
                    },
                    {
                        'Code': '18CS52',
                        'Name': 'Computer Networks',
                        'Internal': '20',
                        'External': '58',
                        'Total': '78',
                        'Result': 'P' 
                    },
                    {
                        'Code': '18CS53',
                        'Name': 'Database Management Systems',
                        'Internal': '19',
                        'External': '62',
                        'Total': '81',
                        'Result': 'P'
                    },
                    {
                        'Code': '18CS54',
                        'Name': 'Automata Theory and Computability',
                        'Internal': '17',
                        'External': '54',
                        'Total': '71',
                        'Result': 'P' if i % 5 != 0 else 'F'
                    },
                    {
                        'Code': '18CS55',
                        'Name': 'Application Development using Python',
                        'Internal': '19',
                        'External': '57',
                        'Total': '76',
                        'Result': 'P'
                    }
                ]
            }
            sample_results.append(student)
        
        # Create a timestamp for the filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_filename = f"demo_results_{timestamp}.xlsx"
        
        # Create sample logs
        logs = [
            f"Processing demo data for {len(sample_results)} students",
            "This is sample data for demonstration purposes only",
            "No actual scraping is performed in demo mode"
        ]
        
        return jsonify({
            'status': 'success',
            'message': f'Successfully generated {len(sample_results)} demo results',
            'data': sample_results,
            'logs': logs,
            'filename': excel_filename,
            'demo': True
        })
        
    except Exception as e:
        return jsonify({
            'status': 'error',
            'message': f"Error generating demo data: {str(e)}"
        }), 500

@app.route('/api/run_script', methods=['POST'])
def run_script():
    """Handle the run_script endpoint for compatibility with the original UI."""
    try:
        data = request.json
        start_usn = data.get('start_usn')
        end_usn = data.get('end_usn')
        interactive_mode = data.get('interactive_mode', False)
        
        if not start_usn or not end_usn:
            return jsonify({'error': 'Start and end USN required'}), 400
            
        # If we're in demo mode or Selenium is not available, use demo data instead
        if os.environ.get('FORCE_DEMO') == 'True':
            # Extract USN pattern and number
            # Support for flexible USN formats like 1AT20CS001, 1AT21IS002, etc.
            usn_pattern_match = re.match(r'(\d+[A-Z]{2}\d{2}[A-Z]{2})(\d{3})', start_usn)
            
            if usn_pattern_match:
                usn_prefix = usn_pattern_match.group(1)  # e.g., 1AT22CS
                start_num = int(usn_pattern_match.group(2))  # e.g., 001 as int
            else:
                try:
                    start_num = int(start_usn)
                    usn_prefix = "1AT22CS"  # Default prefix if only numbers provided
                except ValueError:
                    return jsonify({'error': 'Invalid start USN format'}), 400
            
            end_pattern_match = re.match(r'(\d+[A-Z]{2}\d{2}[A-Z]{2})(\d{3})', end_usn)
            
            if end_pattern_match:
                end_prefix = end_pattern_match.group(1)
                end_num = int(end_pattern_match.group(2))
                
                # Validate that prefixes match
                if end_prefix != usn_prefix:
                    return jsonify({'error': 'Start and end USN must have the same college, year, and branch codes'}), 400
            else:
                try:
                    end_num = int(end_usn)
                except ValueError:
                    return jsonify({'error': 'Invalid end USN format'}), 400
                    
            # Create demo results
            sample_results = []
            for i in range(start_num, end_num + 1):
                usn = f"{usn_prefix}{str(i).zfill(3)}"
                student = {
                    'USN': usn,
                    'Name': f"Demo Student {i}",
                    'Semester': "5",
                    'Total': str(350 + i * 10),
                    'Result': 'PASS' if i % 5 != 0 else 'FAIL',
                    'Subjects': []
                }
                sample_results.append(student)
            
            # Create a timestamp for the filename
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            excel_filename = f"demo_results_{timestamp}.xlsx"
            
            return jsonify({
                'status': 'success',
                'message': f'Generated {len(sample_results)} sample results (Demo Mode)',
                'data': sample_results,
                'filename': excel_filename
            })
        else:
            # Call the scrape function directly with the parameters
            try:
                # Setup driver
                if interactive_mode:
                    os.environ['MANUAL_CAPTCHA'] = 'True'
                
                driver = setup_driver()
                if driver is None:
                    return jsonify({
                        'status': 'error',
                        'message': 'Failed to initialize Chrome WebDriver. Please check if Chrome is installed correctly.'
                    }), 500
                
                # Parse USN numbers
                # Support for flexible USN formats
                usn_pattern_match = re.match(r'(\d+[A-Z]{2}\d{2}[A-Z]{2})(\d{3})', start_usn)
                
                if usn_pattern_match:
                    usn_prefix = usn_pattern_match.group(1)  # e.g., 1AT22CS
                    start_num = int(usn_pattern_match.group(2))  # e.g., 001 as int
                else:
                    try:
                        start_num = int(start_usn)
                        usn_prefix = "1AT22CS"  # Default prefix if only numbers provided
                    except ValueError:
                        return jsonify({'error': 'Invalid start USN format'}), 400
                
                end_pattern_match = re.match(r'(\d+[A-Z]{2}\d{2}[A-Z]{2})(\d{3})', end_usn)
                
                if end_pattern_match:
                    end_prefix = end_pattern_match.group(1)
                    end_num = int(end_pattern_match.group(2))
                    
                    # Validate that prefixes match for batch processing
                    if end_prefix != usn_prefix:
                        return jsonify({'error': 'Start and end USN must have the same college, year, and branch codes for batch processing'}), 400
                else:
                    try:
                        end_num = int(end_usn)
                    except ValueError:
                        return jsonify({'error': 'Invalid end USN format'}), 400
                
                # Validate range
                if start_num > end_num:
                    start_num, end_num = end_num, start_num
                
                # Limit range for server-side processing to avoid timeouts
                if end_num - start_num > 10:
                    end_num = start_num + 10
                
                # Generate USN list with the correct prefix
                usn_list = [f"{usn_prefix}{str(i).zfill(3)}" for i in range(start_num, end_num + 1)]
                
                # Check if we should use manual mode (no automatic CAPTCHA solving)
                manual_mode = interactive_mode or not API_KEY or API_KEY == "YOUR_2CAPTCHA_API_KEY"
                
                try:
                    # Process results
                    all_results, logs = process_results(driver, usn_list, manual_mode)
                    
                    if all_results:
                        # Save results to Excel
                        excel_filename = save_to_excel(all_results)
                        
                        if excel_filename and os.path.exists(excel_filename):
                            # Create a simplified version of results for the response
                            simplified_results = []
                            for result in all_results:
                                student = {
                                    'USN': result.get('USN', ''),
                                    'Name': result.get('Student Name', ''),
                                    'Semester': result.get('Semester', ''),
                                    'Subjects': []
                                }
                                
                                # Add subjects
                                total_marks = 0
                                for key, value in result.items():
                                    if key not in ['USN', 'Student Name', 'Semester']:
                                        if isinstance(value, dict):
                                            subject = {
                                                'Code': key,
                                                'Name': value.get('Subject Name', ''),
                                                'Internal': value.get('Internal', ''),
                                                'External': value.get('External', ''),
                                                'Total': value.get('Total', ''),
                                                'Result': value.get('Result', '')
                                            }
                                            student['Subjects'].append(subject)
                                            
                                            # Calculate total marks
                                            try:
                                                total_marks += int(value.get('Total', '0'))
                                            except ValueError:
                                                pass
                                
                                student['Total'] = str(total_marks)
                                student['Result'] = 'PASS' if all(s.get('Result', '') == 'P' for s in student['Subjects']) else 'FAIL'
                                simplified_results.append(student)
                            
                            return jsonify({
                                'status': 'success',
                                'message': f'Successfully scraped {len(simplified_results)} results',
                                'data': simplified_results,
                                'logs': logs,
                                'filename': excel_filename
                            })
                        else:
                            return jsonify({
                                'status': 'error',
                                'message': 'Failed to save results to Excel file',
                                'logs': logs
                            }), 500
                    else:
                        return jsonify({
                            'status': 'error',
                            'message': 'No results found',
                            'logs': logs
                        }), 400
                finally:
                    # Close the driver
                    try:
                        if driver:
                            driver.quit()
                            print("Driver closed successfully")
                    except Exception as e:
                        print(f"Error closing driver: {str(e)}")
            except Exception as e:
                error_details = traceback.format_exc()
                print(f"Error in run_script: {str(e)}")
                print(error_details)
                return jsonify({
                    'status': 'error',
                    'message': f'An unexpected error occurred: {str(e)}',
                    'details': error_details
                }), 500
    except Exception as e:
        print(f"Error processing request: {str(e)}")
        return jsonify({
            'status': 'error',
            'message': f'Error processing request: {str(e)}'
        }), 400

# Create template directory and index.html if not exists
def create_template_files():
    """Create template directory and index.html if they don't exist."""
    # Create templates directory if not exists
    if not os.path.exists('templates'):
        os.makedirs('templates')
    
    # Create static directory if not exists
    if not os.path.exists('static'):
        os.makedirs('static')
    
    # Create index.html if not exists
    index_html_path = os.path.join('templates', 'index.html')
    if not os.path.exists(index_html_path):
        with open(index_html_path, 'w', encoding='utf-8') as f:
            f.write('''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>VTU Results Scraper</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body {
            padding-top: 20px;
            padding-bottom: 20px;
            background-color: #f8f9fa;
        }
        .container {
            max-width: 800px;
        }
        .card {
            margin-bottom: 20px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        #results-section {
            display: none;
        }
        #loading {
            display: none;
            margin: 20px 0;
        }
        footer {
            text-align: center;
            margin-top: 30px;
            padding-top: 20px;
            border-top: 1px solid #eee;
            color: #777;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="card">
            <div class="card-header bg-primary text-white">
                <h1 class="h3 mb-0">VTU Results Scraper</h1>
            </div>
            <div class="card-body">
                <p class="lead">Enter the USN range to scrape results from VTU website</p>
                
                <div id="api-key-status" class="alert alert-warning mb-3">
                    Checking 2Captcha API key status...
                </div>
                
                <form id="scraper-form">
                    <div class="row mb-3">
                        <div class="col-md-6">
                            <label for="start-usn" class="form-label">Start USN</label>
                            <input type="text" class="form-control" id="start-usn" placeholder="e.g. 1AT22CS001" required>
                        </div>
                        <div class="col-md-6">
                            <label for="end-usn" class="form-label">End USN</label>
                            <input type="text" class="form-control" id="end-usn" placeholder="e.g. 1AT22CS010" required>
                        </div>
                    </div>
                    
                    <div class="d-grid">
                        <button type="submit" class="btn btn-primary">Scrape Results</button>
                    </div>
                </form>
                
                <div id="loading" class="text-center">
                    <div class="spinner-border text-primary" role="status">
                        <span class="visually-hidden">Loading...</span>
                    </div>
                    <p class="mt-2">Scraping VTU results, please wait...</p>
                    <p class="text-muted">This may take a few minutes depending on the number of USNs</p>
                </div>
                
                <div id="logs-section" class="mt-3" style="display: none;">
                    <h5>Processing Logs:</h5>
                    <div id="logs-content" class="border p-2" style="max-height: 200px; overflow-y: auto; font-family: monospace; font-size: 0.8rem; background-color: #f5f5f5;"></div>
                </div>
                
                <div id="error-message" class="alert alert-danger mt-3" style="display: none;"></div>
            </div>
        </div>
        
        <div id="results-section" class="card">
            <div class="card-header bg-success text-white">
                <h2 class="h4 mb-0">Results</h2>
            </div>
            <div class="card-body">
                <p id="results-count" class="mb-3"></p>
                <div class="table-responsive">
                    <table class="table table-striped" id="results-table">
                        <thead>
                            <tr>
                                <th>USN</th>
                                <th>Name</th>
                                <th>Semester</th>
                                <th>Total Marks</th>
                                <th>Result</th>
                            </tr>
                        </thead>
                        <tbody id="results-body"></tbody>
                    </table>
                </div>
                <div class="d-grid gap-2 d-md-flex justify-content-md-center mt-3">
                    <button id="download-excel" class="btn btn-success">Download Excel</button>
                </div>
            </div>
        </div>
        
        <footer>
            <p>VTU Results Scraper - For educational purposes only</p>
            <p>© {{ current_year }}</p>
            <p><a href="/demo">View Demo (No Selenium Required)</a></p>
        </footer>
    </div>

    <script>
        // Check API key status on page load
        document.addEventListener('DOMContentLoaded', async function() {
            try {
                const response = await fetch('/api/check_api_key');
                const result = await response.json();
                
                const apiKeyStatus = document.getElementById('api-key-status');
                
                if (result.api_key_configured) {
                    apiKeyStatus.className = 'alert alert-success mb-3';
                    apiKeyStatus.innerHTML = '<strong>Automatic CAPTCHA Solving:</strong> Enabled. The system will attempt to solve CAPTCHAs automatically.';
                } else {
                    apiKeyStatus.className = 'alert alert-warning mb-3';
                    apiKeyStatus.innerHTML = '<strong>Automatic CAPTCHA Solving:</strong> Disabled. No valid 2Captcha API key is configured. ' +
                        'The system will have limited functionality as CAPTCHAs cannot be solved automatically.';
                }
            } catch (error) {
                console.error('Error checking API key status:', error);
            }
        });

        document.getElementById('scraper-form').addEventListener('submit', async function(e) {
            e.preventDefault();
            
            const startUsn = document.getElementById('start-usn').value;
            const endUsn = document.getElementById('end-usn').value;
            
            document.getElementById('error-message').style.display = 'none';
            document.getElementById('loading').style.display = 'block';
            document.getElementById('results-section').style.display = 'none';
            document.getElementById('logs-section').style.display = 'none';
            
            try {
                const response = await fetch('/api/scrape', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                    },
                    body: JSON.stringify({ start_usn: startUsn, end_usn: endUsn }),
                });
                
                const result = await response.json();
                
                // Display logs if available
                if (result.logs && result.logs.length > 0) {
                    const logsContent = document.getElementById('logs-content');
                    logsContent.innerHTML = result.logs.join('<br>');
                    document.getElementById('logs-section').style.display = 'block';
                }
                
                if (response.ok) {
                    displayResults(result);
                } else {
                    let errorMessage = result.message || 'An error occurred while scraping results';
                    showError(errorMessage);
                }
            } catch (error) {
                showError('Failed to connect to the server. Please try again later.');
                console.error(error);
            } finally {
                document.getElementById('loading').style.display = 'none';
            }
        });
        
        function displayResults(data) {
            const resultsBody = document.getElementById('results-body');
            resultsBody.innerHTML = '';
            
            document.getElementById('results-count').textContent = `Found ${data.data.length} results`;
            
            data.data.forEach(student => {
                const row = document.createElement('tr');
                
                const usnCell = document.createElement('td');
                usnCell.textContent = student.USN;
                row.appendChild(usnCell);
                
                const nameCell = document.createElement('td');
                nameCell.textContent = student.Name;
                row.appendChild(nameCell);
                
                const semesterCell = document.createElement('td');
                semesterCell.textContent = student.Semester;
                row.appendChild(semesterCell);
                
                const totalCell = document.createElement('td');
                totalCell.textContent = student.Total;
                row.appendChild(totalCell);
                
                const resultCell = document.createElement('td');
                resultCell.textContent = student.Result;
                resultCell.className = student.Result === 'PASS' ? 'text-success fw-bold' : 'text-danger fw-bold';
                row.appendChild(resultCell);
                
                resultsBody.appendChild(row);
            });
            
            document.getElementById('results-section').style.display = 'block';
            
            // Setup download button
            document.getElementById('download-excel').onclick = function() {
                window.location.href = '/download/' + data.filename;
            };
        }
        
        function showError(message) {
            const errorElement = document.getElementById('error-message');
            errorElement.textContent = message;
            errorElement.style.display = 'block';
        }
    </script>
    
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>''')
        print(f"Created {index_html_path}")
    
    # Create demo.html
    demo_html_path = os.path.join('templates', 'demo.html')
    if not os.path.exists(demo_html_path):
        with open(demo_html_path, 'w', encoding='utf-8') as f:
            f.write('''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>VTU Results Scraper - Demo Mode</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body {
            padding-top: 20px;
            padding-bottom: 20px;
            background-color: #f8f9fa;
        }
        .container {
            max-width: 800px;
        }
        .card {
            margin-bottom: 20px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        #results-section {
            display: none;
        }
        #loading {
            display: none;
            margin: 20px 0;
        }
        footer {
            text-align: center;
            margin-top: 30px;
            padding-top: 20px;
            border-top: 1px solid #eee;
            color: #777;
        }
        .demo-badge {
            margin-left: 10px;
            font-size: 0.7em;
            vertical-align: middle;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="card">
            <div class="card-header bg-info text-white">
                <h1 class="h3 mb-0">VTU Results Scraper <span class="badge bg-warning demo-badge">DEMO MODE</span></h1>
            </div>
            <div class="card-body">
                <p class="lead">Demo Mode - No Selenium Required</p>
                
                <div class="alert alert-info mb-3">
                    <p><strong>Demo Mode Information:</strong></p>
                    <p>This is a demonstration of the VTU Results Scraper without using Selenium. It generates sample data to show how the application works.</p>
                    <p>No actual scraping is performed, and the data shown is not real.</p>
                </div>
                
                <form id="demo-form">
                    <div class="row mb-3">
                        <div class="col-md-6">
                            <label for="start-usn" class="form-label">Start USN Number</label>
                            <input type="number" class="form-control" id="start-usn" min="1" max="120" value="1">
                        </div>
                        <div class="col-md-6">
                            <label for="end-usn" class="form-label">End USN Number</label>
                            <input type="number" class="form-control" id="end-usn" min="1" max="120" value="5">
                        </div>
                    </div>
                    
                    <div class="d-grid">
                        <button type="submit" class="btn btn-info">Generate Sample Results</button>
                    </div>
                </form>
                
                <div id="loading" class="text-center">
                    <div class="spinner-border text-info" role="status">
                        <span class="visually-hidden">Loading...</span>
                    </div>
                    <p class="mt-2">Generating sample results...</p>
                </div>
                
                <div id="logs-section" class="mt-3" style="display: none;">
                    <h5>Processing Logs:</h5>
                    <div id="logs-content" class="border p-2" style="max-height: 200px; overflow-y: auto; font-family: monospace; font-size: 0.8rem; background-color: #f5f5f5;"></div>
                </div>
                
                <div id="error-message" class="alert alert-danger mt-3" style="display: none;"></div>
            </div>
        </div>
        
        <div id="results-section" class="card">
            <div class="card-header bg-success text-white">
                <h2 class="h4 mb-0">Demo Results</h2>
            </div>
            <div class="card-body">
                <p id="results-count" class="mb-3"></p>
                <div class="table-responsive">
                    <table class="table table-striped" id="results-table">
                        <thead>
                            <tr>
                                <th>USN</th>
                                <th>Name</th>
                                <th>Semester</th>
                                <th>Total Marks</th>
                                <th>Result</th>
                            </tr>
                        </thead>
                        <tbody id="results-body"></tbody>
                    </table>
                </div>
                <div class="mt-3">
                    <p class="text-muted text-center"><em>Note: These are sample results for demonstration purposes only.</em></p>
                </div>
            </div>
        </div>
        
        <div class="d-grid gap-2 d-md-flex justify-content-md-start mt-3">
            <a href="/" class="btn btn-outline-primary">Back to Main Scraper</a>
        </div>
        
        <footer>
            <p>VTU Results Scraper - Demo Mode</p>
            <p>© {{ current_year }}</p>
        </footer>
    </div>

    <script>
        document.getElementById('demo-form').addEventListener('submit', async function(e) {
            e.preventDefault();
            
            const startUsn = document.getElementById('start-usn').value;
            const endUsn = document.getElementById('end-usn').value;
            
            // Validate input
            if (parseInt(endUsn) - parseInt(startUsn) > 20) {
                showError('Please limit your range to 20 USNs maximum for demo purposes.');
                return;
            }
            
            document.getElementById('error-message').style.display = 'none';
            document.getElementById('loading').style.display = 'block';
            document.getElementById('results-section').style.display = 'none';
            document.getElementById('logs-section').style.display = 'none';
            
            try {
                const response = await fetch('/api/demo', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                    },
                    body: JSON.stringify({ start_usn: startUsn, end_usn: endUsn }),
                });
                
                const result = await response.json();
                
                // Display logs if available
                if (result.logs && result.logs.length > 0) {
                    const logsContent = document.getElementById('logs-content');
                    logsContent.innerHTML = result.logs.join('<br>');
                    document.getElementById('logs-section').style.display = 'block';
                }
                
                if (response.ok) {
                    displayResults(result);
                } else {
                    let errorMessage = result.message || 'An error occurred while generating demo results';
                    showError(errorMessage);
                }
            } catch (error) {
                showError('Failed to connect to the server. Please try again later.');
                console.error(error);
            } finally {
                document.getElementById('loading').style.display = 'none';
            }
        });
        
        function displayResults(data) {
            const resultsBody = document.getElementById('results-body');
            resultsBody.innerHTML = '';
            
            document.getElementById('results-count').textContent = `Generated ${data.data.length} sample results`;
            
            data.data.forEach(student => {
                const row = document.createElement('tr');
                
                const usnCell = document.createElement('td');
                usnCell.textContent = student.USN;
                row.appendChild(usnCell);
                
                const nameCell = document.createElement('td');
                nameCell.textContent = student.Name;
                row.appendChild(nameCell);
                
                const semesterCell = document.createElement('td');
                semesterCell.textContent = student.Semester;
                row.appendChild(semesterCell);
                
                const totalCell = document.createElement('td');
                totalCell.textContent = student.Total;
                row.appendChild(totalCell);
                
                const resultCell = document.createElement('td');
                resultCell.textContent = student.Result;
                resultCell.className = student.Result === 'PASS' ? 'text-success fw-bold' : 'text-danger fw-bold';
                row.appendChild(resultCell);
                
                resultsBody.appendChild(row);
            });
            
            document.getElementById('results-section').style.display = 'block';
        }
        
        function showError(message) {
            const errorElement = document.getElementById('error-message');
            errorElement.textContent = message;
            errorElement.style.display = 'block';
        }
    </script>
    
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>''')
        print(f"Created {demo_html_path}")

if __name__ == '__main__':
    # Create template files if they don't exist
    create_template_files()
    
    # Get port from environment or use default
    port = int(os.environ.get('PORT', 5000))
    
    # Print startup message
    print(f"\n{'='*50}")
    print("VTU Results Scraper Web Application")
    print(f"{'='*50}")
    
    # Set up development mode for better manual CAPTCHA entry
    if not os.environ.get('FORCE_DEMO') and not os.environ.get('DEVELOPMENT'):
        print("Setting DEVELOPMENT mode for better manual CAPTCHA entry")
        os.environ['DEVELOPMENT'] = 'True'
    
    # Test Selenium availability
    try:
        print("Testing Selenium availability...")
        test_driver = setup_driver()
        if test_driver:
            test_driver.quit()
            print("✓ Selenium is working correctly")
        else:
            print("✗ Selenium driver setup failed")
            print("Forcing demo mode as Selenium is not available")
            os.environ['FORCE_DEMO'] = 'True'
    except Exception as e:
        print(f"✗ Selenium test failed: {str(e)}")
        print("Forcing demo mode as Selenium is not available")
        os.environ['FORCE_DEMO'] = 'True'
    
    print(f"Starting server on port {port}")
    print(f"Visit http://localhost:{port} in your browser")
    if os.environ.get('FORCE_DEMO') == 'True':
        print("NOTICE: Running in DEMO MODE only")
    else:
        print("NOTICE: Running in INTERACTIVE MODE - you'll need to type the CAPTCHAs manually")
        print("A Chrome browser window will open for each USN. Type the CAPTCHA and wait for results to load.")
    print(f"{'='*50}")
    
    # Check 2Captcha API key
    if API_KEY == "20480f95adb6216bc0e788f58c343c11":
        print("\nWARNING: The provided 2Captcha API key appears to be invalid.")
        print("Automatic CAPTCHA solving will be disabled.")
        print("You will need to enter CAPTCHAs manually.")
        
        # Temporarily disable the API key
        API_KEY = ""
    
    # Run the app
    app.run(host='0.0.0.0', port=port, debug=os.environ.get('DEVELOPMENT', False)) 