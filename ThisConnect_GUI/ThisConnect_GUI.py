# 
# UPDATED: Handle merged StudentID-Name format in Connect Portal
# - Portal changed from separate "StudentId" and "Student Name" fields to single "Student" field
# - New format: "StudentID-Name" (e.g., "12345678-John Doe") 
# - Excel still has separate ID and Name columns, code transforms them for portal matching
# - Updated placeholder from 'StudentId' to 'Student'
# - Enhanced matching logic with multiple strategies for reliability
#
import tkinter as tk
from tkinter import ttk  # Import ttk
from tkinter import filedialog, messagebox, simpledialog
import threading
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.firefox.service import Service as FirefoxService
from selenium.webdriver.edge.service import Service as EdgeService
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.firefox import GeckoDriverManager
from webdriver_manager.microsoft import EdgeChromiumDriverManager
import os
import datetime
import queue
import time  # Add time module for delays
import subprocess  # Add subprocess for opening files

# Global variable declaration
text_messages = None # Declare text_messages as a global variable
gradesheet_url = None # Declare gradesheet_url as a global variable
url_entered_event = threading.Event() # Declare url_entered_event as a global variable
is_browser_open = False # Declare is_browser_open as a global variable to check if the browser is open
driver = None  # Declare driver as a global variabl
continue_execution = True  # Global flag to control execution = True # Global flag to control thread execution
open_browsers = []  # List to track open browser instances

# Function definitions (create_log_file, init_driver, process_gradesheet)

def create_log_file(excel_base_name, unmatched_students, portal_only_students=None):
    # Get the directory where the script is located
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # Create the log file name with the current date and time
    current_time = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    log_file_name = f"mismatch_log_{excel_base_name}_{current_time}.txt"

    # Construct the full path for the log file in the script directory
    log_file_path = os.path.join(script_dir, log_file_name)

    # Check if the log file already exists
    file_mode = 'a' if os.path.exists(log_file_path) else 'w'

    with open(log_file_path, file_mode) as log_file:
        if file_mode == 'w':
            # Write the header if the file is new
            log_file.write("Mismatch Student Log\n")
            log_file.write("Time: " + current_time + "\n\n")
            
        # Write unmatched student details (in Excel but not on portal)
        if unmatched_students:
            log_file.write("=== STUDENTS IN EXCEL BUT NOT ON PORTAL ===\n")
            log_file.write("SL #, ID #, Name\n")
            for student in unmatched_students:
                log_file.write(f"{student['sl']}, {student['id']}, {student['name']}\n")
            log_file.write("\n")
                
        # Write portal-only student details (on portal but not in Excel)
        if portal_only_students:
            log_file.write("=== STUDENTS ON PORTAL BUT NOT IN EXCEL ===\n")
            log_file.write("ID #, Name\n")
            for student in portal_only_students:
                log_file.write(f"{student['id']}, {student['name']}\n")

    print(f"Mismatch information logged in '{log_file_path}'.")
    return log_file_path  # Return the path of the log file

# Function to open the log file in the default text editor
def open_log_file(log_file_path):
    try:
        os.startfile(log_file_path)  # This works on Windows
    except AttributeError:
        # For MacOS and Linux, you can use subprocess
        subprocess.call(['open', log_file_path])  # MacOS
        subprocess.call(['xdg-open', log_file_path])  # Linux

def open_log_folder(folder_path):
    try:
        os.startfile(folder_path)  # This works on Windows
    except AttributeError:
        subprocess.call(['open', folder_path])  # MacOS
        subprocess.call(['xdg-open', folder_path])  # Linux

# Function to initialize the WebDriver based on user choice
def init_driver(browser_choice):
    global open_browsers

    # Open the new browser based on the choice
    try:
        new_driver = None
        if browser_choice.lower() == 'chrome':
            new_driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
        elif browser_choice.lower() == 'firefox':
            new_driver = webdriver.Firefox(service=FirefoxService(GeckoDriverManager().install()))
        elif browser_choice.lower() == 'edge':
            new_driver = webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()))

        if new_driver:
            open_browsers.append(new_driver)  # Add the new driver to the list

        return new_driver
    except Exception as e:
        messagebox.showerror("Error", f"Failed to open {browser_choice}. Is it installed?")
        return None

    
def on_main_window_close():
    global is_browser_open, driver, open_browsers

    if open_browsers or is_browser_open or driver:
        if messagebox.askokcancel("Quit", "Close all browsers and exit?"):
            for browser in open_browsers:
                try:
                    browser.quit()
                except Exception as e:
                    print(f"Error closing browser: {e}")
            open_browsers.clear()
            is_browser_open = False
        else:
            # User chose not to close the application
            return
    root.quit()
    root.destroy()

def check_browser_status():
    global is_browser_open, driver
    try:
        if driver and driver.title:  # Attempt to get the title of the current page
            root.after(1000, check_browser_status)  # Re-check after 1 second
    except:
        is_browser_open = False
        print("Browser has been closed by the user.")
        # root.destroy()

def is_student_data_visible():
    """Check if student data is visible on the current page"""
    try:
        # Check for student fields with new merged format
        student_count = driver.execute_script('return document.querySelectorAll("input[placeholder=\'Student\']").length')
        return student_count > 0
    except Exception as e:
        print(f"Error checking student data: {e}")
        return False

# Create a queue for inter-thread communication
url_queue = queue.Queue()

def show_url_prompt():
    url_prompt_message = (
        "Please provide the URL from the current Connect Portal browser tab:\n\n"
        "1. Make sure you are already on the mark entry page\n"
        "2. Make sure the student list has been loaded\n"
        "3. Copy the URL from the browser address bar and paste it below\n\n"
        "IMPORTANT: Please DO NOT interact (click, resize, move, etc.) "
        "with the browser window while processing is ongoing."
    )
    url = simpledialog.askstring("Input", url_prompt_message)
    if url is None:
        # User clicked Cancel or closed the prompt
        url_queue.put(None)  # Signal that the operation was canceled
        return None
    elif valid_url(url):
        url_queue.put(url)  # Put the valid URL into the queue
        return url
    else:
        # The valid_url function now shows an error message
        return show_url_prompt()  # Call the function again for a valid URL

def valid_url(url):
    if not url:
        return False
        
    expected_url_part = "connect.bracu.ac.bd/app/exam-controller/mark-entry/final"
    
    # Check if the URL is properly formed and contains the expected part
    is_valid = url.startswith("https://") and expected_url_part in url
    
    if not is_valid:
        messagebox.showwarning(
            "URL Validation",
            "The URL should be from the Connect Portal mark entry page.\n\n"
            "Example: https://connect.bracu.ac.bd/app/exam-controller/mark-entry/final/...\n\n"
            "Please make sure you've navigated to the mark entry page and loaded the student list."
        )
    
    return is_valid


def start_processing_thread():
    processing_thread = threading.Thread(target=process_gradesheet)
    label_status.config(text="", foreground="black")
    clear_text_messages()
    processing_thread.start()

def update_text_messages(message, tag=None):
    global text_messages
    
    # Create text_messages widget if it doesn't exist yet
    if text_messages is None:
        # Create the text_messages widget if it doesn't exist
        label_messages = tk.Label(root, text="Processing Messages:")
        label_messages.pack()

        text_messages = tk.Text(root, height=10, width=50, font=("Helvetica", 10))
        scrollbar = tk.Scrollbar(root, command=text_messages.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text_messages.config(yscrollcommand=scrollbar.set)
        text_messages.pack()

        # Configure a tag for red and bold text
        text_messages.tag_configure("unmatched", foreground="red", font=("Helvetica", 10, "bold"))
    
    text_messages.insert(tk.END, message + "\n", tag)
    text_messages.see(tk.END)  # Auto-scroll to the end

def clear_file_path():
    entry_file_path.delete(0, tk.END)

def clear_text_messages():
    global text_messages
    if text_messages is not None:
        text_messages.delete('1.0', tk.END)

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        # Check if the file extension is for Excel files
        if file_path.lower().endswith(('.xlsx', '.xls')):
            entry_file_path.delete(0, tk.END)
            entry_file_path.insert(0, file_path)
            label_status.config(text="", foreground="black")
            clear_text_messages()
        else:
            messagebox.showerror("Error", "The selected file is not an Excel file.")

def enable_gui_elements():
    button_browse.config(state='normal')
    button_clear.config(state='normal')
    button_process.config(state='normal', text='Start Processing')
    entry_file_path.config(state='normal')
    browser_dropdown.config(state='readonly')
    total_marks_entry.config(state='normal')

def disable_gui_elements():
    button_browse.config(state='disabled')
    button_clear.config(state='disabled')
    button_process.config(state='disabled', text='Processing started, Please wait...')
    entry_file_path.config(state='disabled')
    browser_dropdown.config(state='disabled')
    total_marks_entry.config(state='disabled')

def process_gradesheet():
    global driver, is_browser_open, gradesheet_url, text_messages, continue_thread_execution

    # Get values from GUI elements
    excel_path = entry_file_path.get()
    print(f"Selected File: {excel_path}")  # Debug print
    browser_choice = browser_option.get()
    print(f"Selected Browser: {browser_choice}")  # Debug print
    total_marks = total_marks_var.get() or "100"  # Get the total marks value or default to 100

    if not excel_path or not browser_choice:
        messagebox.showerror("Error", "Please specify the gradesheet file and browser choice.")
        enable_gui_elements()
        return

    # Check if the file exists
    if not os.path.exists(excel_path):
        messagebox.showerror("Error", "The specified file does not exist. Please check the path and try again.")
        enable_gui_elements()
        return

    # Disable GUI elements
    disable_gui_elements()

    # Read the Excel file into a Pandas DataFrame
    df = pd.read_excel(excel_path)

    # Extract the base name of the Excel file (without path and extension)
    excel_base_name = os.path.splitext(os.path.basename(excel_path))[0]

    # Initialize WebDriver
    driver = init_driver(browser_choice)
    if driver is None:
        # Re-enable GUI elements and return
        enable_gui_elements()
        return
    check_browser_status()

    # Open the login page
    login_url = 'https://connect.bracu.ac.bd/'  # Connect Portal login URL
    
    try:
        driver.get(login_url)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load login page: {str(e)}\n\nPlease check your internet connection and try again.")
        driver.quit()
        enable_gui_elements()
        return

    # Manual login message
    messagebox.showinfo("Login", "Please log in to Connect Portal, navigate to the Mark Entry page and Load the student list. \n\nThen click OK to continue.")

    # Now schedule the URL prompt to show after login confirmation
    root.after(100, show_url_prompt)

    # Wait for the URL from the prompt
    gradesheet_url = url_queue.get()
    if gradesheet_url is None:
        # User canceled the operation
        enable_gui_elements()
        return
    elif gradesheet_url:
        # Initialize text message area if needed
        update_text_messages(f"Processing gradesheet at: {gradesheet_url}")
        
        # Check if we're already on the correct page
        current_url = driver.current_url
        
        # If we're not on the correct URL or the URL has changed, navigate to it
        if current_url != gradesheet_url:
            update_text_messages(f"Navigating to the gradesheet page...")
            try:
                driver.get(gradesheet_url)
                # Add a delay to let the page fully load
                time.sleep(5)  # Adjust the delay as needed
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load gradesheet page: {str(e)}")
                enable_gui_elements()
                return
        else:
            update_text_messages(f"Already on the correct page, proceeding...")
            # Refresh the page only if student data isn't visible
            if not is_student_data_visible():
                update_text_messages("No student data found. Refreshing the page...")
                driver.refresh()
                time.sleep(5)  # Wait for the page to refresh
        
        # Check if we can find student data on the page
        if not is_student_data_visible():
            update_text_messages("No student data found on the page. Waiting for data to load...")
            # Ask user to ensure data is loaded
            messagebox.showinfo(
                "Page Loading", 
                "No student data detected on the page.\n\n"
                "Please ensure you:\n"
                "1. Are on the correct mark entry page\n"
                "2. Have loaded all student data\n"
                "3. Can see the student list on the page\n\n"
                "Click OK once the data is visible."
            )
            
            # Check again after user confirmation
            if not is_student_data_visible():
                messagebox.showerror(
                    "Error", 
                    "Still cannot detect student data. Please try:\n"
                    "1. Refreshing the page\n"
                    "2. Reloading the student list\n"
                    "3. Restarting the process"
                )
                enable_gui_elements()
                return
    else:
        messagebox.showerror("Error", "No URL entered. Exiting.")
        driver.quit()
        return
    
    # Clear the text_messages widget if it exists
    if text_messages is not None:
        text_messages.delete('1.0', tk.END)

    unmatched_students = [] # List to store unmatched students

    # Get all student data from the page using JavaScript
    update_text_messages("Collecting student data from the page...")
    js_script = """
    const rows = document.querySelectorAll("input[placeholder='Student']");
    const studentData = [];
    
    // Also get available status options from the first dropdown
    let availableStatuses = [];
    const firstStatusSelect = document.querySelector("mat-select");
    if (firstStatusSelect) {
        // Click to open dropdown temporarily to see options
        firstStatusSelect.click();
        
        // Wait briefly for options to appear
        setTimeout(() => {
            const options = document.querySelectorAll('mat-option');
            availableStatuses = Array.from(options).map(opt => opt.textContent.trim());
            // Close dropdown
            document.body.click();
        }, 300);
    }

    rows.forEach((studentInput, index) => {
      const statusSelect = document.querySelectorAll("mat-select")[index];
      
      // Parse the combined "StudentID-Name" field
      const studentValue = studentInput?.value || '';
      let studentId = '';
      let studentName = '';
      
      if (studentValue.includes('-')) {
        const parts = studentValue.split('-');
        studentId = parts[0].trim();
        studentName = parts.slice(1).join('-').trim(); // Handle names with hyphens
      } else {
        // Fallback: treat entire value as ID if no hyphen found
        studentId = studentValue.trim();
        studentName = 'Unknown';
      }
      
      studentData.push({
        id: studentId,
        name: studentName,
        combinedValue: studentValue, // Store original combined value for reference
        index: index,  // Store the index for later use
        currentStatus: statusSelect?.textContent?.trim() || 'Present'  // Get current status
      });
    });

    return {studentData: studentData, availableStatuses: availableStatuses};
    """
    
    # Wait a moment for the dropdown detection
    time.sleep(1)
    result = driver.execute_script(js_script)
    portal_students = result.get('studentData', [])
    available_statuses = result.get('availableStatuses', [])
    
    if available_statuses:
        update_text_messages(f"Available status options on portal: {available_statuses}")
    else:
        update_text_messages("Could not detect available status options")
    
    # Check if we need to set the Total Marks field (usually 100)
    update_text_messages("Checking for Total Marks field...")
    set_total_marks = f"""
    const totalMarksInput = document.querySelector("input[placeholder='Total marks']");
    if (totalMarksInput) {{
        totalMarksInput.value = '{total_marks}';
        totalMarksInput.dispatchEvent(new Event('input', {{ bubbles: true }}));
        totalMarksInput.dispatchEvent(new Event('change', {{ bubbles: true }}));
        return true;
    }}
    return false;
    """
    
    total_marks_set = driver.execute_script(set_total_marks)
    if total_marks_set:
        update_text_messages(f"Total Marks field set to {total_marks}")
    else:
        update_text_messages("Total Marks field not found")
    
    if not portal_students:
        messagebox.showerror("Error", "Could not find student data on the page. Please ensure you're on the correct page.")
        enable_gui_elements()
        return

    update_text_messages(f"Found {len(portal_students)} students on the current page")
    
    # Create a dictionary for quick lookup by student ID
    portal_students_dict = {student['id']: student for student in portal_students}
    
    # Also create a lookup by combined format for better matching
    portal_combined_dict = {}
    for student in portal_students:
        if student['id'] and student['name']:
            combined_key = f"{student['id']}-{student['name']}"
            portal_combined_dict[combined_key] = student
    
    # Create a set of Excel student IDs for quick lookups
    excel_student_ids = set()
    
    # Check if Status column exists in Excel
    has_status_column = 'Status' in df.columns
    if has_status_column:
        update_text_messages("✅ Status column found in Excel - will update both grades and attendance status")
        update_text_messages("⏱️  Note: Status updates will increase processing time")
    else:
        update_text_messages("ℹ️  No 'Status' column found in Excel - will only update grades (faster processing)")
    
    # Prepare a list to track which students need grade updates
    students_to_update = []
    unmatched_students = [] # Students in Excel but not on portal
    portal_only_students = [] # Students on portal but not in Excel
    status_summary = {}  # Track status distribution for debugging
    
    for index, excel_row in df.iterrows():
        excel_student_id = str(excel_row['ID #']).strip()  # Get the student ID from Excel
        excel_grade = "{:.2f}".format(round(excel_row['Total'], 2))  # Round and format the grade
        excel_sl = str(excel_row['Sl #']).strip()  # Get the serial number
        excel_name = str(excel_row['Name']).strip()  # Get the student name
        
        # Get the status from Excel (Present/Absent/On-Hold)
        excel_status = "Present"  # Default status
        if has_status_column:
            if pd.notna(excel_row['Status']):  # Check if the status is not NaN
                excel_status = str(excel_row['Status']).strip()
                # Normalize the status value
                if excel_status.lower() in ["present", "p"]:
                    excel_status = "Present"
                elif excel_status.lower() in ["absent", "a"]:
                    excel_status = "Absent"
                elif excel_status.lower() in ["on-hold", "onhold", "on hold", "oh", "on_hold"]:
                    excel_status = "On-Hold"
                else:
                    update_text_messages(f"Warning: Unrecognized status '{excel_status}' for student {excel_student_id}, defaulting to Present")
                    excel_status = "Present"  # Default if unrecognized
            else:
                update_text_messages(f"No status found for student {excel_student_id}, defaulting to Present")
        
        # Track status distribution for debugging (only if status column exists)
        if has_status_column:
            if excel_status in status_summary:
                status_summary[excel_status] += 1
            else:
                status_summary[excel_status] = 1
        
        # Add to our set of Excel student IDs
        excel_student_ids.add(excel_student_id)

        # Create combined format for portal matching (Excel ID + Name)
        excel_combined = f"{excel_student_id}-{excel_name}"
        
        # Try multiple matching strategies
        matched_student = None
        match_method = ""
        
        # Strategy 1: Direct ID match
        if excel_student_id in portal_students_dict:
            matched_student = portal_students_dict[excel_student_id]
            match_method = "direct_id"
        
        # Strategy 2: Combined format match
        elif excel_combined in portal_combined_dict:
            matched_student = portal_combined_dict[excel_combined]
            match_method = "combined_format"
        
        # Strategy 3: Fuzzy matching (case-insensitive, whitespace tolerant)
        else:
            for portal_student in portal_students:
                portal_combined = f"{portal_student['id']}-{portal_student['name']}"
                
                # Case-insensitive comparison
                if excel_combined.lower() == portal_combined.lower():
                    matched_student = portal_student
                    match_method = "fuzzy_case"
                    break
                
                # ID-only fallback match
                if excel_student_id.lower() == portal_student['id'].lower():
                    matched_student = portal_student
                    match_method = "fuzzy_id"
                    break

        if matched_student:
            if has_status_column:
                update_text_messages(f"✓ Match found ({match_method}): Excel[{excel_student_id}-{excel_name}] → Portal[{matched_student['id']}-{matched_student['name']}] - Status: {excel_status}, Grade: {excel_grade}")
            else:
                update_text_messages(f"✓ Match found ({match_method}): Excel[{excel_student_id}-{excel_name}] → Portal[{matched_student['id']}-{matched_student['name']}] - Grade: {excel_grade}")
            
            # Add to the list of students to update (ALL matched students should be updated)
            students_to_update.append({
                'index': matched_student['index'],
                'id': excel_student_id,
                'grade': excel_grade,
                'status': excel_status if has_status_column else None,
                'name': excel_name,
                'portal_id': matched_student['id'],  # Store portal ID for verification
                'portal_name': matched_student['name']  # Store portal name for verification
            })
        else:
            unmatched_students.append({'sl': excel_sl, 'id': excel_student_id, 'name': excel_name})
            update_text_messages(f"✗ No match found: Excel[{excel_student_id}-{excel_name}]", "unmatched")
    
    # Find students on portal but not in Excel
    for student in portal_students:
        if student['id'] not in excel_student_ids:
            portal_only_students.append(student)
            update_text_messages(f"Student ID {student['id']} on portal but not in Excel.", "unmatched")
    
    # Display status summary for debugging (only if status column exists)
    if status_summary and has_status_column:
        update_text_messages(f"Excel status distribution: {status_summary}")        # Validate that our student indexing is correct before proceeding
        if students_to_update:
            update_text_messages("Validating student index mapping...")
            js_validate_indices = """
            const studentInputs = document.querySelectorAll("input[placeholder='Student']");
            const validation = [];
            
            studentInputs.forEach((input, index) => {
                const studentValue = input.value || '';
                let studentId = '';
                let studentName = '';
                
                if (studentValue.includes('-')) {
                    const parts = studentValue.split('-');
                    studentId = parts[0].trim();
                    studentName = parts.slice(1).join('-').trim();
                } else {
                    studentId = studentValue.trim();
                    studentName = 'Unknown';
                }
                
                validation.push({
                    index: index,
                    id: studentId,
                    name: studentName,
                    combinedValue: studentValue
                });
            });
            
            return validation;
            """
        
        portal_validation = driver.execute_script(js_validate_indices)            # Check if our students_to_update indices match the portal
        index_mismatches = []
        for student in students_to_update:
            if student['index'] < len(portal_validation):
                portal_student = portal_validation[student['index']]
                # Use portal_id for comparison since that's what we'll find on the portal
                if portal_student['id'] != student['portal_id']:
                    index_mismatches.append({
                        'excel_id': student['id'],
                        'excel_name': student['name'],
                        'portal_id': portal_student['id'],
                        'portal_name': portal_student['name'],
                        'expected_portal_id': student['portal_id'],
                        'index': student['index']
                    })
        if index_mismatches:
            update_text_messages(f"⚠ WARNING: Found {len(index_mismatches)} index mismatches!")
            for mismatch in index_mismatches[:5]:  # Show first 5 mismatches
                update_text_messages(f"  Index {mismatch['index']}: Excel expects portal ID {mismatch['expected_portal_id']}, found {mismatch['portal_id']}")
            
            # Re-map the students based on actual portal order
            update_text_messages("Re-mapping student indices based on portal order...")
            portal_dict = {student['id']: i for i, student in enumerate(portal_validation)}
            
            corrected_students = []
            for student in students_to_update:
                # Use portal_id for lookup since that's what should be on the portal
                if student['portal_id'] in portal_dict:
                    student['index'] = portal_dict[student['portal_id']]
                    corrected_students.append(student)
                    update_text_messages(f"Corrected index for Excel ID {student['id']} (Portal ID {student['portal_id']}): now at index {student['index']}")
                else:
                    update_text_messages(f"⚠ Portal ID {student['portal_id']} for Excel ID {student['id']} not found in current portal list")
            
            students_to_update = corrected_students
        else:
            update_text_messages("✓ All student indices validated successfully")
    
    # Update all the matched students at once to avoid individual page reloads
    if students_to_update:
        if has_status_column:
            update_text_messages(f"Preparing to update grades and status for {len(students_to_update)} students...")
        else:
            update_text_messages(f"Preparing to update grades for {len(students_to_update)} students...")
            update_text_messages("ℹ️  Status updates skipped - no 'Status' column in Excel")
        
        # CRITICAL: Validate student indices BEFORE any updates
        update_text_messages("🔍 Validating student indices to prevent wrong student updates...")
        js_validate_all = """
        const studentInputs = document.querySelectorAll("input[placeholder='Student']");
        const validation = [];
        
        studentInputs.forEach((input, index) => {
            const studentValue = input.value.trim() || '';
            let studentId = '';
            let studentName = '';
            
            if (studentValue.includes('-')) {
                const parts = studentValue.split('-');
                studentId = parts[0].trim();
                studentName = parts.slice(1).join('-').trim();
            } else {
                studentId = studentValue.trim();
                studentName = 'Unknown';
            }
            
            validation.push({
                index: index,
                id: studentId,
                name: studentName,
                combinedValue: studentValue
            });
        });
        
        return validation;
        """
        
        portal_validation = driver.execute_script(js_validate_all)
        
        # Check for index mismatches and correct them BEFORE any updates
        corrected_students = []
        index_mismatches = 0
        
        for student in students_to_update:
            # Find the correct index for this student's portal ID
            correct_index = None
            for i, portal_student in enumerate(portal_validation):
                if portal_student['id'] == student['portal_id']:
                    correct_index = i
                    break
            
            if correct_index is not None:
                if correct_index != student['index']:
                    index_mismatches += 1
                    update_text_messages(f"⚠ Index corrected for Excel ID {student['id']} (Portal ID {student['portal_id']}): {student['index']} → {correct_index}")
                    student['index'] = correct_index
                
                corrected_students.append(student)
                if has_status_column:
                    update_text_messages(f"✓ Validated: Excel[{student['id']}-{student['name']}] → Portal[{student['portal_id']}-{student['portal_name']}] at index {student['index']} - Grade: {student['grade']}, Status: {student['status']}")
                else:
                    update_text_messages(f"✓ Validated: Excel[{student['id']}-{student['name']}] → Portal[{student['portal_id']}-{student['portal_name']}] at index {student['index']} - Grade: {student['grade']}")
            else:
                update_text_messages(f"✗ Portal ID {student['portal_id']} for Excel ID {student['id']} not found in current portal page - SKIPPING")
        
        students_to_update = corrected_students
        
        if index_mismatches > 0:
            update_text_messages(f"⚠ Corrected {index_mismatches} index mismatches")
        
        if not students_to_update:
            update_text_messages("❌ No valid students to update after validation")
            return
        
        update_text_messages(f"✅ Validation complete. Proceeding with {len(students_to_update)} validated students")
        
        # Now update grades safely - each student individually with verification
        update_text_messages("📝 Starting grade updates...")
        grade_update_count = 0
        
        for student in students_to_update:
            try:
                # Double-check the student ID before updating grades
                js_grade_update = f"""
                const marksInputs = document.querySelectorAll("input[placeholder='Marks']");
                const studentInputs = document.querySelectorAll("input[placeholder='Student']");
                
                const targetInput = marksInputs[{student['index']}];
                const studentInput = studentInputs[{student['index']}];
                
                if (!studentInput) {{
                    return {{
                        success: false,
                        error: 'Student input field not found'
                    }};
                }}
                
                // Extract student ID from combined field
                const studentValue = studentInput.value.trim();
                let actualStudentId = '';
                if (studentValue.includes('-')) {{
                    actualStudentId = studentValue.split('-')[0].trim();
                }} else {{
                    actualStudentId = studentValue.trim();
                }}
                
                // CRITICAL: Verify we're updating the correct student using portal ID
                if (actualStudentId !== '{student['portal_id']}') {{
                    return {{
                        success: false,
                        error: `SAFETY CHECK FAILED: Expected Portal ID '{student['portal_id']}', found '${{actualStudentId}}' at index {student['index']}`
                    }};
                }}
                
                if (!targetInput) {{
                    return {{
                        success: false,
                        error: 'Marks input field not found'
                    }};
                }}
                
                const oldValue = targetInput.value;
                targetInput.value = '{student['grade']}';
                targetInput.dispatchEvent(new Event('input', {{ bubbles: true }}));
                targetInput.dispatchEvent(new Event('change', {{ bubbles: true }}));
                targetInput.dispatchEvent(new Event('blur', {{ bubbles: true }}));
                
                return {{
                    success: true,
                    oldValue: oldValue,
                    newValue: targetInput.value,
                    studentId: actualStudentId,
                    combinedValue: studentValue
                }};
                """
                
                grade_result = driver.execute_script(js_grade_update)
                
                if grade_result.get('success'):
                    grade_update_count += 1
                    update_text_messages(f"✓ Grade updated for Excel[{student['id']}] → Portal[{student['portal_id']}]: {grade_result.get('oldValue', 'empty')} → {student['grade']}")
                else:
                    update_text_messages(f"✗ Grade update FAILED for Excel[{student['id']}] → Portal[{student['portal_id']}]: {grade_result.get('error', 'Unknown error')}")
                    continue  # Skip status update if grade update failed
                
                # Small delay between individual updates
                time.sleep(0.3)
                
            except Exception as e:
                update_text_messages(f"✗ Exception during grade update for Excel[{student['id']}] → Portal[{student['portal_id']}]: {str(e)}")
                continue
        
        update_text_messages(f"📝 Grade updates complete: {grade_update_count}/{len(students_to_update)} successful")
        
        # Initialize status update counter (used in final summary regardless of Status column presence)
        status_update_count = 0
        
        # Only handle status updates if Excel has Status column
        if has_status_column:
            # Now handle status updates - each student individually with full isolation
            update_text_messages("🎯 Starting status updates...")
            
            # Capture current status of all students before updates
            js_capture_all_statuses = """
            const statusSelects = document.querySelectorAll("mat-select");
            const studentInputs = document.querySelectorAll("input[placeholder='Student']");
            const currentStatuses = [];
            
            statusSelects.forEach((select, index) => {
                const studentInput = studentInputs[index];
                const studentValue = studentInput?.value?.trim() || '';
                let studentId = '';
                
                if (studentValue.includes('-')) {
                    studentId = studentValue.split('-')[0].trim();
                } else {
                    studentId = studentValue.trim();
                }
                
                const currentStatus = select?.textContent?.trim() || 'UNKNOWN';
                currentStatuses.push({
                    index: index,
                    id: studentId,
                    status: currentStatus,
                    combinedValue: studentValue
                });
            });
            
            return currentStatuses;
            """
            
            pre_update_statuses = driver.execute_script(js_capture_all_statuses)
            update_text_messages(f"📊 Pre-update status snapshot: {len(pre_update_statuses)} students captured")
            
            update_text_messages("📋 Status update summary:")
            for i, student in enumerate(students_to_update[:5]):  # Show first 5 for debugging
                # Find current status for this student
                current_status = "UNKNOWN"
                for status_info in pre_update_statuses:
                    if status_info['id'] == student['portal_id']:  # Use portal_id for lookup
                        current_status = status_info['status']
                        break
                update_text_messages(f"  {i+1}. Excel[{student['id']}] → Portal[{student['portal_id']}], Current: {current_status} → Target: {student['status']}")
            if len(students_to_update) > 5:
                update_text_messages(f"  ... and {len(students_to_update) - 5} more students")
            
            for student in students_to_update:
                try:
                    # First check if status update is needed
                    js_status_check = f"""
                    const statusSelects = document.querySelectorAll("mat-select");
                    const studentInputs = document.querySelectorAll("input[placeholder='Student']");
                    
                    const statusSelect = statusSelects[{student['index']}];
                    const studentInput = studentInputs[{student['index']}];
                    
                    if (!studentInput) {{
                        return {{
                            needsUpdate: false,
                            error: 'Student input field not found'
                        }};
                    }}
                    
                    // Extract student ID from combined field
                    const studentValue = studentInput.value.trim();
                    let actualStudentId = '';
                    if (studentValue.includes('-')) {{
                        actualStudentId = studentValue.split('-')[0].trim();
                    }} else {{
                        actualStudentId = studentValue.trim();
                    }}
                    
                    // CRITICAL: Verify we're checking the correct student using portal ID
                    if (actualStudentId !== '{student['portal_id']}') {{
                        return {{
                            needsUpdate: false,
                            error: `SAFETY CHECK FAILED: Expected Portal ID '{student['portal_id']}', found '${{actualStudentId}}' at index {student['index']}`
                        }};
                    }}
                    
                    if (!statusSelect) {{
                        return {{
                            needsUpdate: false,
                            error: 'Status select not found'
                        }};
                    }}
                    
                    const currentStatus = statusSelect.textContent.trim();
                    const targetStatus = '{student['status']}';
                    
                    return {{
                        needsUpdate: currentStatus !== targetStatus,
                        currentStatus: currentStatus,
                        targetStatus: targetStatus,
                        studentId: actualStudentId,
                        combinedValue: studentValue
                    }};
                        """
                    
                    status_info = driver.execute_script(js_status_check)
                    
                    if status_info.get('error'):
                        update_text_messages(f"✗ Status check error for Excel[{student['id']}] → Portal[{student['portal_id']}]: {status_info['error']}")
                        continue
                    
                    if not status_info.get('needsUpdate'):
                        update_text_messages(f"- Status already correct for Excel[{student['id']}] → Portal[{student['portal_id']}]: {student['status']}")
                        continue
                    
                    update_text_messages(f"🎯 Updating status for Excel[{student['id']}] → Portal[{student['portal_id']}]: {status_info['currentStatus']} → {student['status']}")
                    
                    # Open dropdown for this specific student
                    js_open_dropdown = f"""
                    const statusSelects = document.querySelectorAll("mat-select");
                    const statusSelect = statusSelects[{student['index']}];
                    
                    if (statusSelect) {{
                        statusSelect.click();
                        return true;
                    }}
                    return false;
                    """
                    
                    dropdown_opened = driver.execute_script(js_open_dropdown)
                    
                    if not dropdown_opened:
                        update_text_messages(f"✗ Failed to open dropdown for Excel[{student['id']}] → Portal[{student['portal_id']}]")
                        continue
                    
                    # Wait for dropdown to open
                    time.sleep(1.2)
                    
                    # First ensure all dropdowns are closed to prevent interference
                    driver.execute_script("document.body.click();")
                    time.sleep(0.5)
                    
                    # Re-open our specific dropdown
                    js_reopen_dropdown = f"""
                    const statusSelects = document.querySelectorAll("mat-select");
                    const ourDropdown = statusSelects[{student['index']}];
                    if (ourDropdown) {{
                        ourDropdown.click();
                        return true;
                    }}
                    return false;
                    """
                    
                    reopened = driver.execute_script(js_reopen_dropdown)
                    if not reopened:
                        update_text_messages(f"✗ Failed to reopen dropdown for Excel[{student['id']}] → Portal[{student['portal_id']}]")
                        continue
                
                    # Wait for options to appear
                    time.sleep(1.0)
                    
                    # Select the correct option with enhanced matching
                    js_select_option = f"""
                    const targetStatus = '{student['status']}';
                    const options = document.querySelectorAll('mat-option');
                    const availableOptions = Array.from(options).map(opt => opt.textContent.trim());
                    
                    // Try exact match
                    for (const option of options) {{
                        const optionText = option.textContent.trim();
                        if (optionText === targetStatus) {{
                            option.click();
                            return {{success: true, matched: optionText, method: 'exact', availableOptions: availableOptions}};
                        }}
                    }}
                    
                    // Try case-insensitive match
                    for (const option of options) {{
                        const optionText = option.textContent.trim();
                        if (optionText.toLowerCase() === targetStatus.toLowerCase()) {{
                            option.click();
                            return {{success: true, matched: optionText, method: 'case-insensitive', availableOptions: availableOptions}};
                        }}
                    }}
                    
                    // Try partial matching for common variations
                    const targetLower = targetStatus.toLowerCase();
                    for (const option of options) {{
                        const optionText = option.textContent.trim();
                        const optionLower = optionText.toLowerCase();
                        
                        // Enhanced On-Hold matching
                        if (targetLower.includes('hold') || targetLower.includes('on-hold') || targetLower === 'on-hold') {{
                            if (optionLower.includes('hold') || optionLower.includes('on-hold') || 
                                optionLower.includes('onhold') || optionLower.includes('on_hold') ||
                                optionLower.includes('on hold') || optionLower === 'oh') {{
                                option.click();
                                return {{success: true, matched: optionText, method: 'partial-hold', availableOptions: availableOptions}};
                            }}
                        }}
                        
                        // Absent matching
                        if (targetLower.includes('absent') || targetLower === 'absent') {{
                            if (optionLower.includes('absent') || optionLower.includes('abs')) {{
                                option.click();
                                return {{success: true, matched: optionText, method: 'partial-absent', availableOptions: availableOptions}};
                            }}
                        }}
                        
                        // Present matching
                        if (targetLower.includes('present') || targetLower === 'present') {{
                            if (optionLower.includes('present') || optionLower.includes('pres')) {{
                                option.click();
                                return {{success: true, matched: optionText, method: 'partial-present', availableOptions: availableOptions}};
                            }}
                        }}
                    }}
                    
                    // Close dropdown if no match found
                    document.body.click();
                    return {{
                        success: false,
                        availableOptions: availableOptions,
                        targetStatus: targetStatus
                    }};
                    """
                
                    option_result = driver.execute_script(js_select_option)
                    
                    # Wait for the selection to complete
                    time.sleep(1.0)
                    
                    # Verify the status was actually set correctly
                    js_verify_status = f"""
                    const statusSelects = document.querySelectorAll("mat-select");
                    const statusSelect = statusSelects[{student['index']}];
                    
                    if (statusSelect) {{
                        return {{
                            actualStatus: statusSelect.textContent.trim(),
                            targetStatus: '{student['status']}'
                        }};
                    }}
                    return {{error: 'Status select not found for verification'}};
                    """
                    
                    verification = driver.execute_script(js_verify_status)
                    
                    if option_result.get('success'):
                        # Double-check that the status was actually set
                        if verification.get('actualStatus') == student['status']:
                            status_update_count += 1
                            update_text_messages(f"✓ Status VERIFIED for Excel[{student['id']}] → Portal[{student['portal_id']}]: '{verification['actualStatus']}' (method: {option_result['method']})")
                        else:
                            update_text_messages(f"⚠ Status mismatch for Excel[{student['id']}] → Portal[{student['portal_id']}]: Expected '{student['status']}', got '{verification.get('actualStatus', 'UNKNOWN')}'")
                            
                            # Attempt one retry for status correction
                            update_text_messages(f"🔄 Retrying status update for Excel[{student['id']}] → Portal[{student['portal_id']}]...")
                            
                            # Close any open dropdowns
                            driver.execute_script("document.body.click();")
                            time.sleep(0.8)
                            
                            # Try once more
                            retry_opened = driver.execute_script(f"""
                            const statusSelects = document.querySelectorAll("mat-select");
                            const ourDropdown = statusSelects[{student['index']}];
                            if (ourDropdown) {{
                                ourDropdown.click();
                                return true;
                            }}
                            return false;
                            """)
                            
                            if retry_opened:
                                time.sleep(1.0)
                                retry_result = driver.execute_script(js_select_option)
                                time.sleep(1.0)
                                
                                # Verify again
                                final_verification = driver.execute_script(js_verify_status)
                                if final_verification.get('actualStatus') == student['status']:
                                    status_update_count += 1
                                    update_text_messages(f"✓ Status CORRECTED on retry for Excel[{student['id']}] → Portal[{student['portal_id']}]: '{final_verification['actualStatus']}'")
                                else:
                                    update_text_messages(f"✗ Status retry failed for Excel[{student['id']}] → Portal[{student['portal_id']}]: Still showing '{final_verification.get('actualStatus', 'UNKNOWN')}'")
                            else:
                                update_text_messages(f"✗ Failed to reopen dropdown for retry on Excel[{student['id']}] → Portal[{student['portal_id']}]")
                    else:
                        update_text_messages(f"✗ Status update failed for Excel[{student['id']}] → Portal[{student['portal_id']}]. Target: '{student['status']}', Available: {option_result.get('availableOptions', [])}")
                    
                    # Ensure any dropdowns are closed before moving to next student
                    driver.execute_script("document.body.click();")
                    time.sleep(0.5)
                    
                    # Longer delay between status updates to ensure complete isolation
                    time.sleep(1.5)
                    
                except Exception as e:
                    update_text_messages(f"✗ Exception during status update for Excel[{student['id']}] → Portal[{student['portal_id']}]: {str(e)}")
            
            update_text_messages(f"🎯 Status updates complete: {status_update_count}/{len(students_to_update)} successful")
            
            # Capture final status of all students to detect any cross-contamination
            update_text_messages("🔍 Performing final status verification...")
            post_update_statuses = driver.execute_script(js_capture_all_statuses)
            
            # Compare pre and post statuses to detect unexpected changes
            unexpected_changes = []
            for pre_status in pre_update_statuses:
                # Find corresponding post status
                post_status = None
                for post in post_update_statuses:
                    if post['id'] == pre_status['id'] and post['index'] == pre_status['index']:
                        post_status = post
                        break
                
                if post_status:
                    # Check if this student was supposed to be updated (use portal_id for lookup)
                    was_target = any(s['portal_id'] == pre_status['id'] for s in students_to_update)
                    
                    if not was_target and pre_status['status'] != post_status['status']:
                        # This student wasn't supposed to change but did
                        unexpected_changes.append({
                            'id': pre_status['id'],
                            'index': pre_status['index'],
                            'before': pre_status['status'],
                            'after': post_status['status']
                        })
            
            if unexpected_changes:
                update_text_messages(f"⚠ WARNING: {len(unexpected_changes)} unexpected status changes detected!")
                for change in unexpected_changes[:3]:  # Show first 3
                    update_text_messages(f"  Student {change['id']} (index {change['index']}): {change['before']} → {change['after']} (NOT EXPECTED)")
            else:
                update_text_messages("✅ No unexpected status changes detected")
        
        # Final summary message that handles both scenarios
        if has_status_column:
            update_text_messages(f"✅ All updates finished: {grade_update_count} grades, {status_update_count} statuses")
        else:
            update_text_messages(f"✅ All updates finished: {grade_update_count} grades (status updates skipped - no Status column)")

    # # Check for unmatched students and create a log file if necessary
    # if len(unmatched_students) > 0:
    #     create_log_file(excel_base_name, unmatched_students)

    # After processing students, create a log file if there are any mismatch students
    if unmatched_students or portal_only_students:
        log_file_path = create_log_file(excel_base_name, unmatched_students, portal_only_students)
        
        # Create message based on what type of mismatches we found
        message = ""
        if unmatched_students and portal_only_students:
            message = f"{len(unmatched_students)} student(s) in Excel not found on portal and {len(portal_only_students)} student(s) on portal not in Excel."
        elif unmatched_students:
            message = f"{len(unmatched_students)} student(s) in Excel not found on portal."
        else:
            message = f"{len(portal_only_students)} student(s) on portal not in Excel."
            
        user_choice = messagebox.askyesnocancel(
            "Mismatch Students Found",
            f"{message} A mismatch log file has been created.\n"
            "Click 'Yes' to open the log file, 'No' to open the log file folder, or 'Cancel' to dismiss."
        )
        if user_choice:  # User clicked 'Yes'
            open_log_file(log_file_path)
        elif user_choice is False:  # User clicked 'No'
            open_log_folder(os.path.dirname(log_file_path))

    # Update the status label in the main thread at the end of processing
    root.after(0, lambda: label_status.config(text="Processing complete. Review Grades before submit", foreground="green", font=("Helvetica", 10, "bold")))

    # Re-enable GUI elements
    enable_gui_elements()

# GUI setup
root = tk.Tk()
root.title("ThisConnect Tool")

# Set the window icon
import os
script_dir = os.path.dirname(os.path.abspath(__file__))
icon_path = os.path.join(script_dir, 'icon.png')
try:
    root.iconphoto(False, tk.PhotoImage(file=icon_path))
except tk.TclError as e:
    print(f"Could not load icon: {e}")
    # Continue without an icon

# Set the size of the main window (width x height)
root.geometry("750x500")  # Increased size for better visibility

# File path entry
label_file_path = tk.Label(root, text="Gradesheet EXCEL File Path:")
label_file_path.pack(pady=(10, 0))
entry_file_path = tk.Entry(root, width=90)
entry_file_path.pack(pady=5)

# Frame for buttons
button_frame = tk.Frame(root)
button_frame.pack(pady=5)

# Browse button
button_browse = tk.Button(button_frame, text="Browse", command=browse_file)
button_browse.pack(side=tk.LEFT, padx=5)

# Clear Field button
button_clear = tk.Button(button_frame, text="Clear Field", command=clear_file_path)
button_clear.pack(side=tk.LEFT, padx=5)

# Browser choice dropdown
label_browser_choice = tk.Label(root, text="Select Browser:")
label_browser_choice.pack(pady=(2, 0))
browser_option = tk.StringVar(root)
browser_dropdown = ttk.Combobox(root, textvariable=browser_option, state="readonly", values=("Firefox", "Chrome", "Edge"), width=20)
browser_dropdown.current(0)  # Set default selection to Firefox
browser_dropdown.pack(pady=2)

# Total marks entry field
total_marks_frame = tk.Frame(root)
total_marks_frame.pack(pady=2)
total_marks_label = tk.Label(total_marks_frame, text="Total Marks:")
total_marks_label.pack(side=tk.LEFT)
total_marks_var = tk.StringVar(root, value="100")  # Default to 100
total_marks_entry = tk.Entry(total_marks_frame, textvariable=total_marks_var, width=5)
total_marks_entry.pack(side=tk.LEFT, padx=5)

# Status update warning label
status_warning_text = ("WARNING: To update final exam status, include a 'Status' column (Present/Absent/On-Hold).\nThis takes longer. Without it, only grades update.")
status_warning_label = tk.Label(
    root,
    text=status_warning_text,
    font=("Helvetica", 9),
    fg="darkblue",
    justify=tk.CENTER,
    wraplength=580
)
status_warning_label.pack(pady=(5, 5))

# Process button
button_process = tk.Button(root, text="Start Processing", command=start_processing_thread)
button_process.pack(pady=(2, 2))  # Add vertical padding below the button

# Status label
label_status = tk.Label(root, text="")
label_status.pack(pady=(2, 2))  # Add vertical padding around the label

# Footer label with Unicode heart emoji
footer_text = u"Built with <3 by FGZ - @theBoringCodes"
footer_label = tk.Label(root, text=footer_text, font=("Helvetica", 10))
footer_label.pack(side=tk.BOTTOM, pady=(5, 5))  # Adjust padding as needed

# Bind the close event
root.protocol("WM_DELETE_WINDOW", on_main_window_close)

# Start the main loop
root.mainloop()
