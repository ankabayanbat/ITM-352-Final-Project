# ITM-352 Final Project - Car Log Uploader with Tkinter GUI
# Enhanced Version with Speed Optimization, Smart Field Mapping, and Detailed Logging
# AUTHOR: Anka Bayanbat

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    InvalidElementStateException,
    WebDriverException,
)
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys

import csv
from datetime import datetime
import time
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from threading import Thread
import traceback

# --- GLOBAL CONFIGURATION ---
WEBSITE_URL = "https://uh.knack.com/travel-log#trip-log-open/"
INPUT_FILE_DEFAULT = "car_log_input.xlsx"
OUTPUT_LOG = "Submission_Log.csv"

AUTHORIZED_USERS = {
    "anka": "anka123",
    "manager": "manager_pass123",
    "guest": "read_only"
}

EXCEL_COLUMNS = [
    'Department', 'Plate', 'Date', 'Start_Time', 'Start_Mileage',
    'End_Time', 'End_Mileage', 'Destination', 'Driver'
]

# --- FORM SELECTORS ---
# Using the robust ancestor:: XPath selectors for connection fields
FORM_SELECTORS = {
    'Department': (By.XPATH, "//label[contains(., 'Department')]/ancestor::div[contains(@class,'kn-input-connection')][1]"),
    'Plate':      (By.XPATH, "//label[contains(., 'Plate') or contains(., 'Vehicle Plate') or contains(., 'Vehicle Plate - Make - Model') or contains(., 'Vehicle')]/ancestor::div[contains(@class,'kn-input-connection')][1]"),
    'Date':       (By.XPATH, "//label[contains(., 'Date')]/following::input[1]"),
    'Start_Time': (By.XPATH, "//label[contains(., 'Start Time')]/following::input[1]"),
    'Start_Mileage': (By.XPATH, "//label[contains(., 'Odometer Start')]/following::input[1]"),
    'End_Time':      (By.XPATH, "//label[contains(., 'End Time')]/following::input[1]"),
    'End_Mileage':   (By.XPATH, "//label[contains(., 'Odometer End')]/following::input[1]"),
    'Destination':   (By.XPATH, "//label[contains(., 'Destination')]/following::input[1]"),
    'Driver': (By.XPATH, "//label[contains(., 'Driver')]/ancestor::div[contains(@class,'kn-input-connection')][1]"),
    'Submit_Button':   (By.CSS_SELECTOR, "button.kn-button.is-primary"),
    'Success_Message': (By.CSS_SELECTOR, ".kn-message.success"),
}

# ============================================================
#                 LOGGING & DATA LOADING
# ============================================================

def initialize_log():
    """Initialize CSV log with detailed column headers for comparison"""
    if not os.path.exists(OUTPUT_LOG):
        with open(OUTPUT_LOG, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            # Log columns now compare EXCEL input vs. ACTUAL selected value
            writer.writerow([
                'Timestamp', 
                'Status', 
                'Error Message', 
                
                'Department (Excel Input)', 
                'Department (Actual Selected)', 
                
                'Plate (Excel Input)', 
                'Plate (Actual Selected)',
                
                'Date', 
                'Start Time', 
                'Start Odometer', 
                'End Time', 
                'End Odometer', 
                'Destination', 
                
                'Driver (Excel Input)',
                'Driver (Actual Selected)'
            ])


def log_submission(trip_data_excel, trip_data_selected, status, error_msg=""):
    """Log submission with EXCEL input and ACTUAL selected field values."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Simple fields are taken from Excel data as they are direct input
    simple_fields = ['Date', 'Start_Time', 'Start_Mileage', 'End_Time', 'End_Mileage', 'Destination']

    with open(OUTPUT_LOG, 'a', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        row = [
            timestamp,
            status,
            error_msg,
            
            # Department Comparison
            trip_data_excel.get('Department', 'N/A'),
            trip_data_selected.get('Department', trip_data_excel.get('Department', 'N/A')), # Default to Excel if not found in Selected
            
            # Plate Comparison
            trip_data_excel.get('Plate', 'N/A'),
            trip_data_selected.get('Plate', trip_data_excel.get('Plate', 'N/A')),
            
            # Simple Fields (taken from Excel)
            *[trip_data_excel.get(field, 'N/A') for field in simple_fields],

            # Driver Comparison
            trip_data_excel.get('Driver', 'N/A'),
            trip_data_selected.get('Driver', trip_data_excel.get('Driver', 'N/A')),
        ]
        writer.writerow(row)


def load_and_clean_data(file_path):
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Input file not found at {file_path}")

    try:
        if file_path.endswith('.xlsx'):
            df = pd.read_excel(file_path)
        else:
            df = pd.read_csv(file_path)
    except Exception as e:
        raise ValueError(f"Error reading file: {e}")

    col_map = {}
    for col_std in EXCEL_COLUMNS:
        col_std_normalized = col_std.lower().replace(' ', '_')
        for col in df.columns:
            col_normalized = str(col).lower().replace(' ', '_')
            if col_normalized == col_std_normalized:
                col_map[col] = col_std
                break

    df.rename(columns=col_map, inplace=True)

    missing_cols = [col for col in EXCEL_COLUMNS if col not in df.columns]
    if missing_cols:
        raise ValueError(f"Missing required columns in input file: {missing_cols}")

    df['Date'] = pd.to_datetime(df['Date'], errors='coerce').dt.strftime('%m/%d/%Y')
    df = df.fillna('')

    return df.to_dict('records')

# ============================================================
#              CONNECTION FIELD HELPER FUNCTIONS
# ============================================================

def scroll_and_click_wrapper(driver, element, field_label):
    """Scroll to element and click - optimized for speed"""
    try:
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
        time.sleep(0.2)
        
        try:
            element.click()
        except Exception:
            driver.execute_script("arguments[0].click();", element)
        
        time.sleep(0.2)
    except Exception as e:
        print(f"   [INFO] Wrapper click for {field_label}: {e}")


def get_connection_input(driver, field_container, field_label):
    """
    Open the Chosen dropdown for this connection field and return
    the search <input> that belongs to THIS field only.
    """
    # If we were given the input directly, just use it
    try:
        if field_container.tag_name.lower() == "input":
            return field_container
    except Exception:
        pass

    # 1) Try to click the Chosen "single" control inside this field
    try:
        toggle = field_container.find_element(By.CSS_SELECTOR, "a.chzn-single")
        try:
            toggle.click()
        except Exception:
            driver.execute_script("arguments[0].click();", toggle)
        time.sleep(0.2)  # give Chosen time to open its dropdown
    except Exception:
        # Fallback: do nothing special; dropdown might already be open
        print(f"   [DEBUG] Could not find chzn-single toggle for {field_label}, using container only")

    # 2) Now look for the search input *inside this same field*
    try:
        input_el = field_container.find_element(By.CSS_SELECTOR, "div.chzn-search input")
        return input_el
    except Exception:
        # Fallback: any input in the container
        try:
            input_el = field_container.find_element(By.XPATH, ".//input")
            return input_el
        except Exception:
            print(f"   [WARN] Could not locate input for {field_label}")
            return None


def type_and_select_connection_option(driver, field_container, input_el, text, field_label):
    """
    Type in the search text and select an option from the Chosen dropdown
    that belongs to THIS field_container. Prioritizes EXACT match.
    
    RETURNS: The text of the selected option, or None if failed.
    """
    try:
        # Clear the field first
        print(f"   [DEBUG] Clearing field before typing '{text}' for {field_label}...")

        try:
            input_el.click()
            time.sleep(0.1)
            input_el.send_keys(Keys.CONTROL + "a")
            input_el.send_keys(Keys.DELETE)
            time.sleep(0.1)
        except Exception:
            pass

        try:
            input_el.clear()
        except Exception:
            pass

        try:
            driver.execute_script("arguments[0].value = '';", input_el)
        except Exception:
            pass

        # Type the search text
        print(f"   [DEBUG] Typing '{text}' into {field_label} search field...")
        input_el.send_keys(text)
        time.sleep(1.2)  # wait for Chosen to filter results

        # Look for results ONLY under this field's container
        print(f"   [DEBUG] Looking for dropdown results for {field_label}...")
        results_lists = field_container.find_elements(By.CSS_SELECTOR, "ul.chzn-results")

        target_text_lower = text.strip().lower()

        for results_list in results_lists:
            try:
                if not results_list.is_displayed():
                    continue

                dropdown_items = results_list.find_elements(By.CSS_SELECTOR, "li.active-result")
                if not dropdown_items:
                    continue

                print(f"   [DEBUG] Found {len(dropdown_items)} options for {field_label}, seeking exact match for '{text}'...")

                best_match = None
                partial_match = None
                first_item = dropdown_items[0] # Store the first item for fallback

                for item in dropdown_items:
                    item_text = item.text.strip()
                    item_text_lower = item_text.lower()
                    
                    if item_text_lower == target_text_lower:
                        # Found the exact match!
                        best_match = item
                        break 
                    elif target_text_lower in item_text_lower and partial_match is None:
                        # Found a partial match, but keep searching for exact
                        partial_match = item 
                
                # Determine which element to click: Exact > Partial > First
                item_to_click = best_match or partial_match or first_item
                
                # --- EXECUTE CLICK ---
                if item_to_click:
                    selected_text = item_to_click.text.strip()
                    match_type = "EXACT" if best_match else ("PARTIAL" if partial_match else "FIRST_FALLBACK")
                    
                    print(f"   [DEBUG] ✓ {match_type} match for {field_label}: '{selected_text}'")
                    driver.execute_script("arguments[0].scrollIntoView({block:'nearest'});", item_to_click)
                    time.sleep(0.2)
                    try:
                        item_to_click.click()
                    except Exception:
                        driver.execute_script("arguments[0].click();", item_to_click)
                    time.sleep(0.7)
                    print(f"   [INFO] ✓ Selected '{selected_text}' for {field_label} ({match_type} match)")
                    # IMPORTANT: Return the selected text
                    return selected_text
                
                # If we reach here, something went wrong, and no dropdown item was clicked
                return None 

            except Exception as e:
                print(f"   [DEBUG] Error processing results list for {field_label}: {e}")
                continue

        # Fallback: basic keyboard navigation (only if NO results list was found/processed)
        print(f"   [DEBUG] No visible dropdown results; using keyboard navigation for {field_label}...")
        input_el.send_keys(Keys.ARROW_DOWN)
        time.sleep(0.3)
        input_el.send_keys(Keys.ENTER)
        time.sleep(0.7)
        # We can't easily capture the selected text here, so we return the input text as a guess
        return text 

    except Exception as e:
        print(f"   [WARN] Could not fill {field_label}: {e}")
        traceback.print_exc()
        return None


def fill_connection_field(driver, field_container, value, field_label):
    """Fill a connection field (Department, Plate, Driver) with smart dropdown handling.
    
    RETURNS: The actual text selected on the webpage, or None if failed.
    """
    text = (str(value) or "").strip()
    if not text:
        print(f"   [WARN] No value for {field_label}")
        return None

    # Log what we’re sending so you can see if something is off
    if field_label == "Plate":
        print(f"   [SENT] {field_label}: '{text}' (license plate)")
    elif field_label == "Driver":
        print(f"   [SENT] {field_label}: '{text}' (driver name)")
    else:
        print(f"   [SENT] {field_label}: '{text}'")

    # Scroll the whole container into view
    scroll_and_click_wrapper(driver, field_container, field_label)

    # Open dropdown and get the search input for THIS field
    input_el = get_connection_input(driver, field_container, field_label)
    if input_el is None:
        return None

    # Type and select within THIS field’s dropdown only
    return type_and_select_connection_option(driver, field_container, input_el, text, field_label)

# ============================================================
#              SUBMIT & SUCCESS HELPER FUNCTIONS
# ============================================================

def find_submit_button(driver):
    """Find submit button in current context or iframes"""
    submit_locators = [
        FORM_SELECTORS['Submit_Button'],
        (By.XPATH, "//button[contains(., 'Submit')]"),
        (By.XPATH, "//span[contains(., 'Submit')]/ancestor::button[1]"),
    ]

    def find_button_here():
        for by, sel in submit_locators:
            try:
                return WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((by, sel))
                )
            except TimeoutException:
                continue
        return None

    button = find_button_here()
    if button is not None:
        return button

    driver.switch_to.default_content()
    frames = driver.find_elements(By.TAG_NAME, 'iframe')
    for i, frame in enumerate(frames):
        try:
            driver.switch_to.frame(frame)
            button = find_button_here()
            if button is not None:
                print(f"   [INFO] Found submit in iframe {i}")
                return button
        except Exception:
            continue

    return None


def click_button_robust(driver, button):
    """Click button with fallback to JS"""
    try:
        button.click()
    except Exception:
        driver.execute_script("arguments[0].click();", button)


def wait_for_success_message(driver):
    """Wait for success message"""
    def has_success_here():
        try:
            WebDriverWait(driver, 3).until(
                EC.presence_of_element_located(FORM_SELECTORS['Success_Message'])
            )
            return True
        except TimeoutException:
            return False

    if has_success_here():
        return True

    driver.switch_to.default_content()
    frames = driver.find_elements(By.TAG_NAME, 'iframe')
    for frame in frames:
        try:
            driver.switch_to.frame(frame)
            if has_success_here():
                return True
        except Exception:
            continue

    return False


def click_submit_and_wait_success(driver):
    """Submit form and wait for success"""
    button = find_submit_button(driver)
    if button is None:
        raise TimeoutException("Could not locate Submit button")

    click_button_robust(driver, button)
    time.sleep(1)  # Brief wait for submission to process
    wait_for_success_message(driver)

# ============================================================
#                 FORM FILLING & SUBMISSION
# ============================================================

def find_form_context(driver):
    """Locate form in main content or iframe"""
    driver.switch_to.default_content()
    by_plate, sel_plate = FORM_SELECTORS['Plate']

    try:
        driver.find_element(by_plate, sel_plate)
        return True
    except Exception:
        pass

    iframes = driver.find_elements(By.TAG_NAME, 'iframe')
    for i, frame in enumerate(frames):
        driver.switch_to.default_content()
        try:
            driver.switch_to.frame(frame)
            if len(driver.find_elements(by_plate, sel_plate)) > 0:
                print(f"   [INFO] Found form in iframe {i}")
                return True
        except Exception:
            continue

    return False


def fill_all_fields_for_trip(driver, trip_data):
    """Fill all fields for one trip
    
    RETURNS: Dictionary of selected values for connection fields (or None if failed).
    """
    print(f"   [DEBUG] Excel data for this trip:")
    print(f"           Department: '{trip_data.get('Department', '')}'")
    print(f"           Plate: '{trip_data.get('Plate', '')}'")
    print(f"           Driver: '{trip_data.get('Driver', '')}'")
    
    excel_to_form_map = {
        'Department': 'Department',
        'Plate': 'Plate',
        'Date': 'Date',
        'Start_Time': 'Start_Time',
        'Start_Mileage': 'Start_Mileage',
        'End_Time': 'End_Time',
        'End_Mileage': 'End_Mileage',
        'Destination': 'Destination',
        'Driver': 'Driver'
    }

    department_filled = False
    selected_values = {} # To store the actual selected text for logging

    for excel_header, selector_key in excel_to_form_map.items():
        value = trip_data.get(excel_header, '')
        if not value:
            print(f"   [SKIP] {selector_key}: No value provided")
            selected_values[excel_header] = 'N/A (Skipped)'
            continue

        by_type, selector = FORM_SELECTORS[selector_key]

        # Special wait: Driver field needs extra wait after Department loads
        if selector_key == 'Driver' and department_filled:
            print(f"   [INFO] Waiting for Driver field to populate (depends on Department)...")
            time.sleep(3)
        
        try:
            # Set a longer timeout for connection fields
            timeout = 10 if selector_key in ['Department', 'Plate', 'Driver'] else 3
            
            element = WebDriverWait(driver, timeout).until(
                EC.presence_of_element_located((by_type, selector))
            )

            if selector_key in ['Department', 'Plate', 'Driver']:
                print(f"   [DEBUG] About to fill connection field {selector_key} with value: '{value}'")
                
                # NEW: Capture the selected text
                selected_text = fill_connection_field(driver, element, value, selector_key)
                
                if selected_text is not None:
                    selected_values[excel_header] = selected_text
                    if selector_key == 'Department':
                        department_filled = True
                        time.sleep(1.5)  # Extra wait after Department for cascading fields
                else:
                    # Connection field failed to select, use original value as fallback
                    selected_values[excel_header] = f"FAILED: Used Input '{value}'"
                    raise Exception(f"Failed to select required value for {selector_key}")

            else:
                # Standard input fields (Date, Time, Mileage, Destination)
                print(f"   [SENT] {selector_key}: '{value}'")
                try:
                    element.clear()
                    element.send_keys(str(value))
                    selected_values[excel_header] = str(value) # Record for consistency
                except InvalidElementStateException:
                    driver.execute_script("arguments[0].value = arguments[1];", element, str(value))
                    selected_values[excel_header] = str(value) # Record for consistency

        except TimeoutException:
            print(f"   [WARN] Could not locate {selector_key} field on page (timeout after {timeout}s)")
            selected_values[excel_header] = f"FAILED: Timeout (Input: {value})"
            return None
        except Exception as e:
            print(f"   [WARN] Error filling {selector_key}: {e}")
            selected_values[excel_header] = f"FAILED: Error (Input: {value})"
            return None

    # For simple fields that didn't go through the above logic (because they were empty), 
    # ensure their input value is logged for comparison
    for key in excel_to_form_map.keys():
        if key not in selected_values:
             selected_values[key] = trip_data.get(key, 'N/A')

    return selected_values


def click_reload_form_button(driver):
    """Click the 'Reload form' button after successful submission"""
    print("   [INFO] Looking for 'Reload form' button...")
    
    reload_locators = [
        (By.XPATH, "//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'reload')]"),
        (By.XPATH, "//a[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'reload')]"),
        (By.XPATH, "//button[contains(., 'Reload')]"),
        (By.XPATH, "//a[contains(., 'Reload')]"),
        (By.XPATH, "//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'new entry')]"),
        (By.XPATH, "//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'submit another')]"),
    ]

    def find_reload_here():
        for by, sel in reload_locators:
            try:
                elems = driver.find_elements(by, sel)
                for el in elems:
                    if el.is_displayed():
                        text = (el.text or "").strip().lower()
                        if "reload" in text or "new" in text or "another" in text:
                            return el
            except Exception:
                continue
        return None

    driver.switch_to.default_content()
    btn = find_reload_here()

    if btn is None:
        frames = driver.find_elements(By.TAG_NAME, 'iframe')
        for i, frame in enumerate(frames):
            try:
                driver.switch_to.frame(frame)
                btn = find_reload_here()
                if btn is not None:
                    print(f"   [INFO] Found 'Reload form' button in iframe {i}")
                    break
            except Exception:
                continue

    if btn is None:
        print("   [WARN] 'Reload form' button not found, will navigate to URL instead")
        return False

    print("   [INFO] Clicking 'Reload form' button...")
    try:
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
        time.sleep(0.3)
        btn.click()
    except Exception:
        try:
            driver.execute_script("arguments[0].click();", btn)
        except Exception as e:
            print(f"   [WARN] Could not click reload button: {e}")
            return False

    time.sleep(2)
    print("   [INFO] Form reloaded successfully")
    return True


def fill_and_submit_trip(driver, trip_data, trip_index, is_last_trip=False):
    """Complete pipeline for one trip with visual feedback"""
    trip_date = trip_data.get('Date', 'N/A')
    print(f"\n--- Processing Trip {trip_index + 1} ({trip_date}) ---")
    
    # Store the results of the form filling to pass to the logger
    selected_values = {}
    success = False
    error_message = ""

    try:
        if not find_form_context(driver):
            raise Exception("Could not locate form fields")

        print("   [INFO] Filling form fields...")
        selected_values = fill_all_fields_for_trip(driver, trip_data)

        if selected_values is None:
            # fill_all_fields_for_trip returned None due to failure
            raise Exception("Form field filling failed. See console warnings.")

        print("   [INFO] ⏸️  Form filled! Pausing 3 seconds for visual verification...")
        time.sleep(3)

        print("   [INFO] Submitting form...")
        click_submit_and_wait_success(driver)

        print("   [SUCCESS] Trip submitted!")
        success = True
        
        if not is_last_trip:
            time.sleep(1)
            if not click_reload_form_button(driver):
                print("   [INFO] Navigating to main page as fallback...")
                driver.switch_to.default_content()
                driver.get(WEBSITE_URL)
                time.sleep(3)
        
        return True

    except Exception as e:
        driver.switch_to.default_content()
        error_message = str(e)
        print(f"   [ERROR] Failed: {error_message}")
        traceback.print_exc()
        
        try:
            driver.get(WEBSITE_URL)
            time.sleep(3)
        except Exception:
            pass
        
        return False

    finally:
        # Log the submission attempt regardless of success
        log_submission(
            trip_data_excel=trip_data, 
            trip_data_selected=selected_values, 
            status=("SUCCESS" if success else "FAILED"),
            error_msg=error_message
        )


# ============================================================
#                       TKINTER GUI
# ============================================================

class CarLogUploader:
    def __init__(self, master):
        self.master = master
        current_os_user = os.environ.get('USERNAME') or os.environ.get('USER') or 'User'
        master.title(f"Car Log Uploader - {current_os_user}")
        master.geometry("700x650")
        master.configure(bg='#f0f0f0')

        self.file_path = tk.StringVar(value=INPUT_FILE_DEFAULT)
        self.status_log = tk.StringVar(value="Ready. Select file and click Upload.")
        self.progress_var = tk.DoubleVar()
        self.driver = None

        self.create_widgets()

    def create_widgets(self):
        # Header
        header = tk.Frame(self.master, bg='#2c3e50', height=60)
        header.pack(fill='x', padx=0, pady=0)
        
        tk.Label(
            header, 
            text="🚗 Car Log Uploader", 
            font=('Segoe UI', 18, 'bold'),
            bg='#2c3e50',
            fg='white'
        ).pack(pady=15)

        # File selection section
        frame1 = tk.LabelFrame(
            self.master, 
            text="📁 Select Input File", 
            padx=20, 
            pady=15,
            font=('Segoe UI', 10, 'bold'),
            bg='#f0f0f0'
        )
        frame1.pack(padx=20, pady=15, fill="x")

        file_frame = tk.Frame(frame1, bg='#f0f0f0')
        file_frame.pack(fill='x')

        tk.Entry(
            file_frame, 
            textvariable=self.file_path, 
            font=('Segoe UI', 10),
            relief=tk.SOLID,
            borderwidth=1
        ).pack(side=tk.LEFT, fill="x", expand=True, ipady=5)
        
        tk.Button(
            file_frame, 
            text="Browse", 
            command=self.browse_file,
            bg='#3498db',
            fg='black',
            font=('Segoe UI', 10),
            relief=tk.FLAT,
            cursor='hand2',
            padx=15
        ).pack(side=tk.LEFT, padx=(10, 0))

        # Upload control section
        frame2 = tk.LabelFrame(
            self.master, 
            text="🚀 Upload Control", 
            padx=20, 
            pady=15,
            font=('Segoe UI', 10, 'bold'),
            bg='#f0f0f0'
        )
        frame2.pack(padx=20, pady=15, fill="x")

        self.upload_button = tk.Button(
            frame2,
            text="▶ Start Upload",
            command=self.start_automation_thread,
            bg='#27ae60',
            fg='black',
            font=('Segoe UI', 12, 'bold'),
            relief=tk.FLAT,
            cursor='hand2',
            height=2
        )
        self.upload_button.pack(fill="x", pady=(0, 15))

        # Progress bar
        style = ttk.Style()
        style.theme_use('clam')
        style.configure(
            "custom.Horizontal.TProgressbar",
            troughcolor='#ecf0f1',
            background='#3498db',
            thickness=25
        )
        
        ttk.Progressbar(
            frame2, 
            variable=self.progress_var, 
            maximum=100,
            mode='determinate',
            style="custom.Horizontal.TProgressbar"
        ).pack(fill="x", pady=(0, 10))

        # Status section
        status_frame = tk.LabelFrame(
            self.master,
            text="📊 Status",
            padx=20,
            pady=15,
            font=('Segoe UI', 10, 'bold'),
            bg='#f0f0f0'
        )
        status_frame.pack(padx=20, pady=(0, 20), fill="both", expand=True)

        status_text = tk.Label(
            status_frame,
            textvariable=self.status_log,
            wraplength=600,
            justify=tk.LEFT,
            font=('Segoe UI', 9),
            bg='white',
            relief=tk.SOLID,
            borderwidth=1,
            padx=10,
            pady=10
        )
        status_text.pack(fill="both", expand=True)

        # Footer
        footer = tk.Frame(self.master, bg='#ecf0f1', height=30)
        footer.pack(fill='x', side='bottom')
        
        tk.Label(
            footer,
            text="ITM-352 Final Project | Anka Bayanbat",
            font=('Segoe UI', 8),
            bg='#ecf0f1',
            fg='#7f8c8d'
        ).pack(pady=5)

    def browse_file(self):
        filename = filedialog.askopenfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.file_path.set(filename)
            self.status_log.set(f"✓ File selected: {os.path.basename(filename)}")

    def start_automation_thread(self):
        self.upload_button.config(
            state=tk.DISABLED,
            text="⏳ Uploading... DO NOT CLOSE",
            bg='#e74c3c'
        )
        self.status_log.set("Starting automation...")
        self.progress_var.set(0)

        t = Thread(target=self.run_automation)
        t.daemon = True
        t.start()

    def safe_set_status(self, text):
        self.master.after(0, self.status_log.set, text)

    def safe_set_progress(self, value):
        self.master.after(0, self.progress_var.set, value)

    def safe_messagebox_info(self, title, message):
        self.master.after(0, messagebox.showinfo, title, message)

    def safe_messagebox_error(self, title, message):
        self.master.after(0, messagebox.showerror, title, message)

    def run_automation(self):
        trip_entries = []
        self.driver = None
        success_count = 0
        total_trips = 0

        try:
            trip_entries = load_and_clean_data(self.file_path.get())
            total_trips = len(trip_entries)
            initialize_log()

            self.safe_set_status("🌐 Opening browser...")

            try:
                self.driver = webdriver.Chrome()
            except WebDriverException as e:
                print(f"[INFO] Using webdriver_manager: {e}")
                service = Service(ChromeDriverManager().install())
                self.driver = webdriver.Chrome(service=service)

            try:
                self.driver.maximize_window()
            except Exception:
                pass

            self.driver.get(WEBSITE_URL)
            time.sleep(3)

            for i, trip in enumerate(trip_entries):
                trip_date = trip.get('Date', 'N/A')
                self.safe_set_status(
                    f"📝 Processing trip {i + 1}/{total_trips}\n"
                    f"Date: {trip_date} | Driver: {trip.get('Driver', 'N/A')}"
                )
                self.safe_set_progress((i / total_trips) * 100)

                is_last = (i == total_trips - 1)
                success = fill_and_submit_trip(self.driver, trip, i, is_last_trip=is_last)
                
                if success:
                    success_count += 1
                else:
                    time.sleep(2)

                self.safe_set_progress(((i + 1) / total_trips) * 100)

            final_message = (
                f"✅ COMPLETE!\n\n"
                f"Successfully submitted: {success_count}/{total_trips} trips\n"
                f"Detailed log saved to: {OUTPUT_LOG}"
            )
            self.safe_set_status(final_message)
            self.safe_messagebox_info("Upload Complete", final_message)

        except Exception as e:
            error_message = f"❌ ERROR: {str(e)}"
            print(error_message)
            traceback.print_exc()
            self.safe_set_status(error_message)
            self.safe_messagebox_error("Error", error_message)

        finally:
            if self.driver:
                try:
                    self.driver.quit()
                except Exception:
                    pass
            
            def reset_button():
                self.upload_button.config(
                    state=tk.NORMAL,
                    text="▶ Start Upload",
                    bg='#27ae60'
                )
            
            self.master.after(0, reset_button)

# ============================================================
#                       LOGIN WINDOW
# ============================================================

class LoginWindow:
    def __init__(self, master):
        self.master = master
        master.title("Login - Car Log Uploader")

        screen_width = master.winfo_screenwidth()
        screen_height = master.winfo_screenheight()
        master.geometry(f"400x350+{int((screen_width / 2) - 200)}+{int((screen_height / 2) - 175)}")
        master.resizable(False, False)
        master.configure(bg='#2c3e50')

        # Header
        header_frame = tk.Frame(master, bg='#34495e', height=80)
        header_frame.pack(fill='x')
        
        tk.Label(
            header_frame,
            text="🔐 Login",
            font=('Segoe UI', 20, 'bold'),
            bg='#34495e',
            fg='white'
        ).pack(pady=25)

        # Login form
        form_frame = tk.Frame(master, bg='#2c3e50')
        form_frame.pack(pady=30, padx=40, fill='both', expand=True)

        tk.Label(
            form_frame,
            text="Username",
            font=('Segoe UI', 10),
            bg='#2c3e50',
            fg='white'
        ).pack(anchor='w', pady=(0, 5))
        
        self.username_entry = tk.Entry(
            form_frame,
            font=('Segoe UI', 11),
            relief=tk.FLAT,
            bg='white'
        )
        self.username_entry.pack(fill='x', ipady=8)
        self.username_entry.bind('<Return>', lambda e: self.handle_login())

        tk.Label(
            form_frame,
            text="Password",
            font=('Segoe UI', 10),
            bg='#2c3e50',
            fg='white'
        ).pack(anchor='w', pady=(15, 5))
        
        self.password_entry = tk.Entry(
            form_frame,
            show="●",
            font=('Segoe UI', 11),
            relief=tk.FLAT,
            bg='white'
        )
        self.password_entry.pack(fill='x', ipady=8)
        self.password_entry.bind('<Return>', lambda e: self.handle_login())

        self.login_button = tk.Button(
            form_frame,
            text="Login",
            command=self.handle_login,
            bg='#27ae60',
            fg='black',
            font=('Segoe UI', 11, 'bold'),
            relief=tk.FLAT,
            cursor='hand2',
            height=2
        )
        self.login_button.pack(fill='x', pady=(20, 0))

        self.error_label = tk.Label(
            form_frame,
            text="",
            fg='#e74c3c',
            bg='#2c3e50',
            font=('Segoe UI', 9)
        )
        self.error_label.pack(pady=(10, 0))

    def handle_login(self):
        user = self.username_entry.get()
        password = self.password_entry.get()

        if user in AUTHORIZED_USERS and AUTHORIZED_USERS[user] == password:
            self.master.destroy()
            root = tk.Tk()
            CarLogUploader(root)
            root.mainloop()
        else:
            self.error_label.config(text="❌ Invalid username or password")
            self.password_entry.delete(0, tk.END)

# ============================================================
#                             MAIN
# ============================================================

if __name__ == "__main__":
    try:
        import pandas as _pd_check
    except ImportError:
        print("Pandas required. Install: pip install pandas")
        sys.exit(1)

    root = tk.Tk()
    LoginWindow(root)
    root.mainloop()