import time
import pandas as pd
from datetime import datetime
import os
import base64
import requests
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# 2Captcha API key - Replace with your actual API key
API_KEY = "20480f95adb6216bc0e788f58c343c11"  # 2Captcha API key

def setup_driver():
    """Set up and return a Chrome WebDriver instance."""
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    
    # Create a new Chrome driver
    driver = webdriver.Chrome(options=chrome_options)
    return driver

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

def process_results(driver, usn_list):
    """Process results for a list of USNs."""
    base_url = "https://results.vtu.ac.in/DJcbcs25/index.php"
    all_results = []
    
    for i, usn in enumerate(usn_list):
        print(f"\nProcessing {usn} ({i+1}/{len(usn_list)})...")
        
        # Navigate to the results page
        driver.get(base_url)
        time.sleep(2)  # Wait for page to load
        
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
                print(f"Found USN input field using selector: {selector}")
                break
            except NoSuchElementException:
                continue
        
        if not usn_input:
            print(f"Could not find USN input field for {usn} using any selector")
            continue
        
        # Enter USN
        usn_input.clear()
        usn_input.send_keys(usn)
        print(f"Entered USN: {usn}")
        
        # Find CAPTCHA input field
        captcha_input = None
        try:
            captcha_input = driver.find_element(By.CSS_SELECTOR, "input[name='captchacode']")
            print("Found CAPTCHA input field")
        except NoSuchElementException:
            print("Could not find CAPTCHA input field")
            continue
        
        # Try to solve CAPTCHA using 2Captcha
        captcha_text = solve_captcha(driver)
        
        if captcha_text:
            print(f"Automatically solved CAPTCHA: {captcha_text}")
            # Enter the CAPTCHA text
            captcha_input.clear()
            captcha_input.send_keys(captcha_text)
            
            # Find and click the submit button
            try:
                submit_button = driver.find_element(By.CSS_SELECTOR, "input[type='submit']")
                submit_button.click()
                print("Clicked submit button")
            except NoSuchElementException:
                print("Could not find submit button")
                continue
        else:
            # Fall back to manual CAPTCHA entry if automatic solving fails
            print("\n" + "-"*50)
            print("Automatic CAPTCHA solving failed. Please complete these steps manually:")
            print("1. Enter the CAPTCHA shown in the image")
            print("2. Click the Submit button")
            print("-"*50)
            
            user_input = input("After submitting, press Enter to continue (or type 'skip'/'exit'): ")
            
            if user_input.lower() == 'skip':
                print(f"Skipping {usn}")
                continue
                
            if user_input.lower() == 'exit':
                print("Exiting the script")
                break
        
        # Wait for results page to load
        print("Waiting for results page to load...")
        time.sleep(3)
        
        # Check if we're still on the input page (CAPTCHA error)
        if "ENTER USN" in driver.page_source or "captchacode" in driver.page_source:
            print(f"Still on input page. CAPTCHA may have been incorrect for {usn}")
            continue
        
        # Extract results
        try:
            # Save the page source for debugging
            with open(f"{usn}_page.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            print(f"Saved page source to {usn}_page.html for debugging")
            
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
                                    print(f"Found USN: {usn_text}")
                                elif "Student Name" in cell_text:
                                    name_text = cells[1].text.strip()
                                    if ":" in name_text:
                                        name_text = name_text.split(":", 1)[1].strip()
                                    student_info["Student Name"] = name_text
                                    print(f"Found Student Name: {name_text}")
                    except Exception as e:
                        print(f"Error processing table row: {str(e)}")
            except Exception as e:
                print(f"Error finding student info table: {str(e)}")
            
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
                            print(f"Found Semester: {student_info['Semester']}")
                            break
                    except Exception as e:
                        print(f"Error processing semester div: {str(e)}")
                
                # If semester is still not found, try one more approach
                if "Semester" not in student_info:
                    # Try to find it in the page source
                    page_source = driver.page_source
                    semester_match = re.search(r'Semester\s*:\s*(\d+)', page_source)
                    if semester_match:
                        student_info["Semester"] = semester_match.group(1)
                        print(f"Found Semester from page source: {student_info['Semester']}")
            except Exception as e:
                print(f"Error finding semester: {str(e)}")
            
            # Initialize results dictionary with student info
            results = {}
            for key, value in student_info.items():
                results[key] = value
            
            # Find the divTable structure that contains the results
            try:
                # First, check if we can find the divTable directly
                div_tables = driver.find_elements(By.CLASS_NAME, "divTable")
                
                if div_tables:
                    print(f"Found {len(div_tables)} divTable elements")
                    
                    for div_table in div_tables:
                        # Check if this is the results table by looking for headers
                        try:
                            header_row = div_table.find_element(By.XPATH, ".//div[contains(@class, 'divTableRow')][1]")
                            header_cells = header_row.find_elements(By.CLASS_NAME, "divTableCell")
                            
                            header_texts = [cell.text.strip() for cell in header_cells]
                            print(f"Found table with headers: {header_texts}")
                            
                            if "Subject Code" in header_texts and "Subject Name" in header_texts:
                                print("This is the results table!")
                                
                                # Get all rows except the header
                                data_rows = div_table.find_elements(By.XPATH, ".//div[contains(@class, 'divTableRow')][position() > 1]")
                                print(f"Found {len(data_rows)} subject rows")
                                
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
                                            
                                            print(f"Found subject: {subject_code} - {subject_name} - Internal: {internal_marks}, External: {external_marks}, Total: {total_marks}, Result: {result}")
                                            
                                            # Store in results dictionary
                                            results[subject_code] = {
                                                'Subject Name': subject_name,
                                                'Internal': internal_marks,
                                                'External': external_marks,
                                                'Total': total_marks,
                                                'Result': result
                                            }
                                    except Exception as e:
                                        print(f"Error processing row: {str(e)}")
                                
                                # We found and processed the results table, so break the loop
                                break
                        except Exception as e:
                            print(f"Error processing table: {str(e)}")
                
                # If we couldn't find results using divTable, try using XPath directly
                if len(results) <= 1:  # Only has USN, no subjects
                    print("Could not find any subject results in the divTable structure")
                    print("Trying alternative approach using XPath directly...")
                    
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
                                    
                                    print(f"Found subject (alt method): {subject_code} - {subject_name} - Internal: {internal_marks}, External: {external_marks}, Total: {total_marks}, Result: {result}")
                                    
                                    # Store in results dictionary
                                    results[subject_code] = {
                                        'Subject Name': subject_name,
                                        'Internal': internal_marks,
                                        'External': external_marks,
                                        'Total': total_marks,
                                        'Result': result
                                    }
                        except Exception as e:
                            print(f"Error processing row (alt method): {str(e)}")
            
            except Exception as e:
                print(f"Error finding or processing divTable: {str(e)}")
            
            if len(results) > 1:  # At least USN and some other data
                all_results.append(results)
                print(f"Successfully extracted results for {usn} with {len(results)-1} subjects")
            else:
                print(f"No results data found for {usn}")
                
        except Exception as e:
            print(f"Error processing results for {usn}: {str(e)}")
    
    return all_results

def save_to_excel(all_results):
    """Save results to Excel file."""
    if not all_results:
        print("\nNo results were collected!")
        return
    
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
    
    # Open the Excel file
    try:
        os.startfile(excel_filename)
        print(f"Excel file opened. Check {excel_filename} for the complete results.")
    except:
        print(f"Excel file saved but could not be opened automatically. Please open {excel_filename} manually.")

def main():
    # Global variable declaration
    global API_KEY
    
    # Ask user for USN range
    print("\nUSN Range Configuration")
    print("----------------------")
    
    # Get starting USN number
    while True:
        start_input = input("Enter starting USN number (e.g., 1 for 1AT22CS001): ").strip()
        
        # Check if input is a full USN
        if start_input.startswith("1AT22CS") and len(start_input) == 10:
            try:
                start_num = int(start_input[7:10])
                break
            except ValueError:
                print("Invalid USN format. Please enter a number or a valid USN.")
        # Check if input is just a number
        elif start_input.isdigit():
            start_num = int(start_input)
            if 1 <= start_num <= 120:
                break
            else:
                print("Number must be between 1 and 120.")
        else:
            print("Please enter a valid number or USN.")
            start_num = 1
            break
    
    # Get ending USN number
    while True:
        end_input = input("Enter ending USN number (e.g., 10 for 1AT22CS010): ").strip()
        
        # Check if input is a full USN
        if end_input.startswith("1AT22CS") and len(end_input) == 10:
            try:
                end_num = int(end_input[7:10])
                break
            except ValueError:
                print("Invalid USN format. Please enter a number or a valid USN.")
        # Check if input is just a number
        elif end_input.isdigit():
            end_num = int(end_input)
            if 1 <= end_num <= 120:
                break
            else:
                print("Number must be between 1 and 120.")
        else:
            print("Please enter a valid number or USN.")
            end_num = min(start_num + 9, 120)  # Default to 10 USNs or up to 120
            break
    
    # Validate range
    if start_num > end_num:
        print(f"Starting number ({start_num}) is greater than ending number ({end_num}). Swapping values.")
        start_num, end_num = end_num, start_num
    
    # Limit range to avoid too many requests
    if end_num - start_num > 20:
        print(f"Range too large ({end_num - start_num + 1} USNs). Limiting to 20 USNs.")
        end_num = start_num + 19
    
    # Generate USN list
    usn_list = [f"1AT22CS{str(i).zfill(3)}" for i in range(start_num, end_num + 1)]
    
    print(f"\nWill process {len(usn_list)} USNs: from {usn_list[0]} to {usn_list[-1]}")
    
    # Check 2Captcha API key
    if API_KEY == "20480f95adb6216bc0e788f58c343c11":
        print("\nWARNING: The provided 2Captcha API key appears to be invalid.")
        print("Automatic CAPTCHA solving will be disabled.")
        print("You will need to enter CAPTCHAs manually.")
        
        # Temporarily disable the API key
        API_KEY = ""
    
    print("\nSetting up the browser...")
    driver = setup_driver()
    
    try:
        print("\n" + "="*50)
        print("INSTRUCTIONS:")
        print("="*50)
        
        if API_KEY and API_KEY != "YOUR_2CAPTCHA_API_KEY":
            print("1. The browser will open to the VTU results page")
            print("2. For each USN, the script will enter the USN automatically")
            print("3. The script will attempt to solve the CAPTCHA automatically using 2Captcha")
            print("4. If automatic CAPTCHA solving fails, you'll be prompted to enter it manually")
            print("5. After the results page loads, the script will extract and save the data")
            print("6. Type 'skip' to skip a USN or 'exit' to stop the process when prompted")
        else:
            print("1. The browser will open to the VTU results page")
            print("2. For each USN, the script will enter the USN automatically")
            print("3. You need to manually enter the CAPTCHA in the browser")
            print("4. Click the Submit button in the browser")
            print("5. After the results page loads, press Enter in the terminal")
            print("6. Type 'skip' to skip a USN or 'exit' to stop the process")
        
        print("="*50)
        print("\nThe script will now start processing USNs...")
        print("="*50)
        
        # Process results
        all_results = process_results(driver, usn_list)
        
        # Save results to Excel
        save_to_excel(all_results)
        
    finally:
        # Close the browser
        print("Closing the browser...")
        driver.quit()

if __name__ == "__main__":
    main() 