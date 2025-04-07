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

class VTUResultsScraper:
    def __init__(self, headless=True, auto_captcha=False, api_key=None, manual_captcha=False):
        """Initialize the VTU results scraper.
        
        Args:
            headless (bool): Whether to run Chrome in headless mode
            auto_captcha (bool): Whether to attempt automatic CAPTCHA solving
            api_key (str): 2Captcha API key for solving CAPTCHAs
            manual_captcha (bool): Whether to wait for manual CAPTCHA input
        """
        self.headless = headless
        self.auto_captcha = auto_captcha
        self.api_key = api_key
        self.manual_captcha = manual_captcha
        self.driver = None
    
    def setup_driver(self):
        """Set up and return a Chrome WebDriver instance."""
        chrome_options = Options()
        
        # If manual CAPTCHA input is enabled, we can't use headless mode
        if self.headless and not self.manual_captcha:
            chrome_options.add_argument("--headless")
            chrome_options.add_argument("--disable-gpu")
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
        
        chrome_options.add_argument("--window-size=1920,1080")
        
        # Create a new Chrome driver
        self.driver = webdriver.Chrome(options=chrome_options)
        return self.driver
    
    def solve_captcha(self, driver):
        """Solve CAPTCHA using 2Captcha service."""
        if not self.auto_captcha or not self.api_key:
            print("Automatic CAPTCHA solving is disabled.")
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
                "key": self.api_key,
                "method": "base64",
                "body": img_base64,
                "json": 1
            }
            
            print("Sending CAPTCHA to 2Captcha service...")
            response = requests.post("https://2captcha.com/in.php", data=data)
            response_json = response.json()
            
            if response_json["status"] == 1:
                request_id = response_json["request"]
                print(f"CAPTCHA sent successfully. Request ID: {request_id}")
                
                # Wait for the CAPTCHA to be solved
                print("Waiting for 2Captcha to solve the CAPTCHA...")
                for attempt in range(30):  # Try for 30 attempts (about 150 seconds)
                    time.sleep(5)
                    get_url = f"https://2captcha.com/res.php?key={self.api_key}&action=get&id={request_id}&json=1"
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
                
                print("Timeout waiting for CAPTCHA solution")
                return None
            else:
                error_msg = response_json["request"]
                print(f"Error sending CAPTCHA to 2Captcha: {error_msg}")
                return None
        except Exception as e:
            print(f"Error preparing CAPTCHA for solving: {str(e)}")
            return None
    
    def extract_subject_marks(self, driver, usn):
        """Extract marks for all subjects from the results page."""
        try:
            # Try to find the results tables
            tables = driver.find_elements(By.CSS_SELECTOR, "table.table")
            
            if not tables or len(tables) < 2:
                print(f"No result tables found for {usn}")
                return {}
            
            # Extract student details
            student_details = {}
            
            try:
                student_table = tables[0]
                rows = student_table.find_elements(By.TAG_NAME, "tr")
                
                for row in rows:
                    cells = row.find_elements(By.TAG_NAME, "td")
                    if len(cells) >= 2:
                        label = cells[0].text.strip().replace(":", "")
                        value = cells[1].text.strip()
                        
                        if label and value:
                            student_details[label] = value
            except Exception as e:
                print(f"Error extracting student details: {str(e)}")
            
            # Extract subjects and marks
            subject_data = {}
            
            try:
                marks_table = tables[1]
                rows = marks_table.find_elements(By.TAG_NAME, "tr")
                
                for row in rows[1:]:  # Skip header row
                    cells = row.find_elements(By.TAG_NAME, "td")
                    
                    if len(cells) >= 7:
                        subject_code = cells[0].text.strip()
                        subject_name = cells[1].text.strip()
                        IA = cells[2].text.strip()
                        external = cells[3].text.strip()
                        total = cells[4].text.strip()
                        result = cells[5].text.strip()
                        
                        if subject_code and subject_name:
                            subject_data[subject_code] = {
                                "Subject Name": subject_name,
                                "Internal Assessment": IA,
                                "External": external,
                                "Total": total,
                                "Result": result
                            }
            except Exception as e:
                print(f"Error extracting subject marks: {str(e)}")
            
            # Extract overall result
            overall_result = {}
            
            try:
                result_div = driver.find_element(By.CSS_SELECTOR, "div.col-md-12 > b.text-bold")
                if result_div:
                    result_text = result_div.text.strip()
                    if "SGPA" in result_text:
                        sgpa = result_text.split("SGPA:")[-1].strip()
                        overall_result["SGPA"] = sgpa
                    
                    # Try to find Total Marks
                    total_div = driver.find_element(By.XPATH, "//div[contains(text(), 'Total Marks:')]")
                    if total_div:
                        total_text = total_div.text.strip()
                        total_marks = total_text.split("Total Marks:")[-1].strip()
                        overall_result["Total Marks"] = total_marks
                    
                    # Find result (Pass/Fail)
                    result_element = driver.find_element(By.XPATH, "//div[contains(text(), 'Result:')]")
                    if result_element:
                        result_text = result_element.text.strip()
                        result = result_text.split("Result:")[-1].strip()
                        overall_result["Result"] = result
            except Exception as e:
                print(f"Error extracting overall result: {str(e)}")
            
            # Combine all data
            result_data = {
                "StudentDetails": student_details,
                "Subjects": subject_data,
                "Overall": overall_result
            }
            
            return result_data
            
        except Exception as e:
            print(f"Error extracting subject marks: {str(e)}")
            return {}
    
    def process_single_usn(self, usn):
        """Process results for a single USN."""
        base_url = "https://results.vtu.ac.in/DJcbcs25/index.php"
        
        try:
            # Navigate to the results page
            self.driver.get(base_url)
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
                    usn_input = self.driver.find_element(By.CSS_SELECTOR, selector)
                    break
                except NoSuchElementException:
                    continue
            
            if not usn_input:
                print(f"Could not find USN input field for {usn}")
                return None
            
            # Enter USN
            usn_input.clear()
            usn_input.send_keys(usn)
            print(f"Entered USN: {usn}")
            
            # Find CAPTCHA input field
            captcha_input = None
            captcha_selectors = [
                "input[placeholder='CAPTCHA CODE']",
                "input[name='captchacode']",
                "input.form-control[type='text']:not([placeholder='ENTER USN'])"
            ]
            
            for selector in captcha_selectors:
                try:
                    captcha_input = self.driver.find_element(By.CSS_SELECTOR, selector)
                    break
                except NoSuchElementException:
                    continue
            
            if not captcha_input:
                print(f"Could not find CAPTCHA input field for {usn}")
                return None
            
            # Handle CAPTCHA input
            if self.manual_captcha:
                # If manual CAPTCHA, wait for user to input
                print(f"Please enter the CAPTCHA for USN {usn} in the browser window and click Submit")
                
                # We'll create a simple alert to indicate the user should enter the CAPTCHA
                self.driver.execute_script(
                    f"alert('Please enter the CAPTCHA for USN {usn} and click Submit. DO NOT close this alert until you have entered the CAPTCHA.');"
                )
                
                # Wait for the alert to be dismissed (user acknowledges)
                try:
                    WebDriverWait(self.driver, 300).until(
                        EC.alert_is_present()
                    )
                    alert = self.driver.switch_to.alert
                    alert.accept()
                except:
                    # Alert might have been closed already
                    pass
                
                # Now we need to wait for the user to submit the form and see results
                initial_url = self.driver.current_url
                
                # Keep checking until either:
                # 1. The URL changes (form submitted and redirected)
                # 2. Tables appear in the page (results loaded)
                # 3. Timeout occurs (10 minutes)
                try:
                    for _ in range(600):  # 600 seconds = 10 minutes
                        # Check if URL has changed or results have loaded
                        current_url = self.driver.current_url
                        if current_url != initial_url or "table" in self.driver.page_source.lower():
                            print("Form submitted and results detected")
                            break
                        
                        # Check for error messages
                        error_msgs = [
                            "invalid usn",
                            "invalid captcha",
                            "wrong captcha",
                            "error",
                            "try again"
                        ]
                        page_source = self.driver.page_source.lower()
                        if any(msg in page_source for msg in error_msgs):
                            print("Error detected. You may need to try again.")
                            # Show an alert to notify the user
                            self.driver.execute_script(
                                "alert('Error detected. The USN or CAPTCHA might be incorrect. Click OK and try again.');"
                            )
                            try:
                                WebDriverWait(self.driver, 60).until(EC.alert_is_present())
                                alert = self.driver.switch_to.alert
                                alert.accept()
                            except:
                                pass
                            return None
                        
                        # Wait a bit before checking again
                        time.sleep(1)
                        
                    # We either found results or timed out
                    # Wait a bit more for the page to fully load
                    time.sleep(5)
                    
                    # Verify that we actually have results
                    if "table" not in self.driver.page_source.lower():
                        print("Timeout or no results found")
                        return None
                        
                except Exception as e:
                    print(f"Error waiting for form submission: {str(e)}")
                    return None
                
            elif self.auto_captcha and self.api_key:
                captcha_text = self.solve_captcha(self.driver)
                
                if captcha_text:
                    captcha_input.clear()
                    captcha_input.send_keys(captcha_text)
                    print(f"Entered CAPTCHA: {captcha_text}")
                else:
                    print("Automatic CAPTCHA solving failed")
                    # Can't continue in headless mode if CAPTCHA solving fails
                    if self.headless:
                        return None
                        
                # Click submit button
                submit_button = None
                button_selectors = [
                    "input[type='submit']",
                    "button[type='submit']",
                    "button.btn-primary",
                    "input.btn-primary"
                ]
                
                for selector in button_selectors:
                    try:
                        submit_button = self.driver.find_element(By.CSS_SELECTOR, selector)
                        break
                    except NoSuchElementException:
                        continue
                
                if not submit_button:
                    print(f"Could not find submit button for {usn}")
                    return None
                
                # Click the submit button
                submit_button.click()
                print(f"Clicked submit button for {usn}")
            else:
                # Can't continue in headless mode without auto CAPTCHA or manual input
                if self.headless:
                    print("Cannot proceed with headless mode without auto CAPTCHA or manual input")
                    return None
            
            # Wait a bit for the page to load
            time.sleep(3)
            
            # Check if we got an error (wrong CAPTCHA)
            error_msgs = [
                "Invalid USN",
                "Invalid captcha",
                "Wrong captcha",
                "Error",
                "try again"
            ]
            
            page_source = self.driver.page_source.lower()
            if any(msg.lower() in page_source for msg in error_msgs):
                print(f"Error submitting form for {usn} - likely wrong CAPTCHA")
                # For manual CAPTCHA, show an alert
                if self.manual_captcha:
                    self.driver.execute_script(
                        "alert('Error detected. Please try again with this USN.');"
                    )
                    try:
                        WebDriverWait(self.driver, 60).until(EC.alert_is_present())
                        alert = self.driver.switch_to.alert
                        alert.accept()
                    except:
                        pass
                return None
            
            # Ask user to confirm the results look good
            if self.manual_captcha:
                try:
                    self.driver.execute_script(
                        "const userConfirm = confirm('Do the results look correct? Click OK to continue, or Cancel to retry this USN.'); return userConfirm;"
                    )
                    # Note: We don't actually need to check the return value - if the user clicks Cancel,
                    # they should manually refresh the page or go back, which will automatically trigger 
                    # the result extraction to fail
                    time.sleep(3)  # Give a moment for any user action
                except:
                    pass
                
            # Extract results
            return self.extract_subject_marks(self.driver, usn)
            
        except Exception as e:
            print(f"Error processing {usn}: {str(e)}")
            return None
    
    def process_usn_range(self, start_usn, end_usn):
        """Process results for a range of USNs.
        
        Args:
            start_usn (str): Starting USN (e.g., 1AT22CS001)
            end_usn (str): Ending USN (e.g., 1AT22CS010)
            
        Returns:
            str: Path to the Excel file containing the results
        """
        try:
            # Validate USN format
            if not start_usn.startswith("1AT22CS") or not end_usn.startswith("1AT22CS"):
                # Extract the numeric part
                try:
                    start_num = int(start_usn)
                    end_num = int(end_usn)
                    start_usn = f"1AT22CS{str(start_num).zfill(3)}"
                    end_usn = f"1AT22CS{str(end_num).zfill(3)}"
                except ValueError:
                    print("Invalid USN format")
                    return None
            
            # Extract numeric parts
            start_num = int(start_usn[7:10])
            end_num = int(end_usn[7:10])
            
            # Validate range
            if start_num > end_num:
                start_num, end_num = end_num, start_num
            
            # Generate USN list
            usn_list = [f"1AT22CS{str(i).zfill(3)}" for i in range(start_num, end_num + 1)]
            
            print(f"Processing {len(usn_list)} USNs: from {usn_list[0]} to {usn_list[-1]}")
            
            # Setup driver
            self.setup_driver()
            
            # Show initial information alert
            if self.manual_captcha:
                try:
                    self.driver.execute_script(
                        f"alert('Starting process for {len(usn_list)} USNs, from {usn_list[0]} to {usn_list[-1]}. Click OK to begin.');"
                    )
                    WebDriverWait(self.driver, 300).until(EC.alert_is_present())
                    alert = self.driver.switch_to.alert
                    alert.accept()
                except:
                    pass
            
            all_results = []
            
            # Process each USN
            for i, usn in enumerate(usn_list):
                print(f"\nProcessing {usn} ({i+1}/{len(usn_list)})...")
                
                # Show progress alert for manual CAPTCHA
                if self.manual_captcha and i > 0:
                    try:
                        self.driver.execute_script(
                            f"alert('Moving to USN {usn} ({i+1} of {len(usn_list)}). Click OK to continue.');"
                        )
                        WebDriverWait(self.driver, 300).until(EC.alert_is_present())
                        alert = self.driver.switch_to.alert
                        alert.accept()
                    except:
                        pass
                
                # For manual CAPTCHA, try only once
                max_attempts = 1 if self.manual_captcha else 3
                
                success = False
                for attempt in range(max_attempts):
                    result = self.process_single_usn(usn)
                    
                    if result:
                        # Prepare data for Excel
                        student_data = {
                            "USN": usn,
                            "Name": result.get("StudentDetails", {}).get("Name", ""),
                            "Semester": result.get("StudentDetails", {}).get("Semester", ""),
                            "Total": result.get("Overall", {}).get("Total Marks", ""),
                            "Result": result.get("Overall", {}).get("Result", ""),
                            "SGPA": result.get("Overall", {}).get("SGPA", "")
                        }
                        
                        # Add subject wise marks
                        subjects = result.get("Subjects", {})
                        for subject_code, subject_data in subjects.items():
                            student_data[f"{subject_code} - {subject_data['Subject Name']}"] = subject_data["Total"]
                        
                        all_results.append(student_data)
                        print(f"Successfully processed {usn}")
                        success = True
                        break
                    else:
                        print(f"Attempt {attempt+1} failed for {usn}")
                
                # If manual CAPTCHA and failed, show an alert
                if self.manual_captcha and not success:
                    try:
                        self.driver.execute_script(
                            f"alert('Failed to process USN {usn}. Moving to the next USN. Click OK to continue.');"
                        )
                        WebDriverWait(self.driver, 300).until(EC.alert_is_present())
                        alert = self.driver.switch_to.alert
                        alert.accept()
                    except:
                        pass
                
                # Add a small delay between requests
                time.sleep(2)
            
            # Show completion alert
            if self.manual_captcha:
                try:
                    self.driver.execute_script(
                        f"alert('All USNs processed! Found results for {len(all_results)} out of {len(usn_list)} USNs. Click OK to finish.');"
                    )
                    WebDriverWait(self.driver, 300).until(EC.alert_is_present())
                    alert = self.driver.switch_to.alert
                    alert.accept()
                except:
                    pass
            
            # Close the driver
            if self.driver:
                self.driver.quit()
                self.driver = None
            
            # Save results to Excel
            if all_results:
                return self.save_to_excel(all_results)
            else:
                print("No results found for any USN")
                return None
            
        except Exception as e:
            print(f"Error processing USN range: {str(e)}")
            # Close the driver
            if self.driver:
                self.driver.quit()
                self.driver = None
            return None
        
    def save_to_excel(self, all_results):
        """Save results to an Excel file.
        
        Args:
            all_results (list): List of dictionaries containing student results
            
        Returns:
            str: Path to the saved Excel file
        """
        if not all_results:
            print("No results to save")
            return None
        
        try:
            # Create a DataFrame
            df = pd.DataFrame(all_results)
            
            # Generate filename with timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            excel_filename = f"vtu_results_{timestamp}.xlsx"
            
            # Save to Excel
            df.to_excel(excel_filename, index=False)
            print(f"Results saved to {excel_filename}")
            
            return excel_filename
            
        except Exception as e:
            print(f"Error saving results to Excel: {str(e)}")
            return None 