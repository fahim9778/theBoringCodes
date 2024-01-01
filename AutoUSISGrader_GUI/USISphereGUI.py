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

# Global variable declaration
text_messages = None # Declare text_messages as a global variable
gradesheet_url = None # Declare gradesheet_url as a global variable
url_entered_event = threading.Event() # Declare url_entered_event as a global variable
is_browser_open = False # Declare is_browser_open as a global variable to check if the browser is open
driver = None  # Declare driver as a global variabl
continue_execution = True  # Global flag to control execution = True # Global flag to control thread execution


# Function definitions (create_log_file, init_driver, process_gradesheet)

def create_log_file(excel_base_name, unmatched_students):
    # Check if there are any unmatched students
    if not unmatched_students:
        return

    # Get the directory where the script is located
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # Create the log file name with the current date and time
    current_time = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    log_file_name = f"unmatched_students_log_{excel_base_name}_{current_time}.txt"

    # Construct the full path for the log file in the script directory
    log_file_path = os.path.join(script_dir, log_file_name)

    # Check if the log file already exists
    file_mode = 'a' if os.path.exists(log_file_path) else 'w'

    with open(log_file_path, file_mode) as log_file:
        if file_mode == 'w':
            # Write the header if the file is new
            log_file.write("Unmatched Student Log\n")
            log_file.write("Time: " + current_time + "\n")
            log_file.write("Student SL in Excel, ID, Name\n")

        # Write unmatched student details
        for student in unmatched_students:
            log_file.write(f"{student['sl']}, {student['id']}, {student['name']}\n")

    print(f"Unmatched students are logged in '{log_file_path}'.")
    return log_file_path  # Return the path of the log file

# Function to open the log file in the default text editor
def open_log_file(log_file_path):
    try:
        os.startfile(log_file_path)  # This works on Windows
    except AttributeError:
        # For MacOS and Linux, you can use subprocess
        import subprocess
        subprocess.call(['open', log_file_path])  # MacOS
        subprocess.call(['xdg-open', log_file_path])  # Linux

def open_log_folder(folder_path):
    try:
        os.startfile(folder_path)  # This works on Windows
    except AttributeError:
        import subprocess
        subprocess.call(['open', folder_path])  # MacOS
        subprocess.call(['xdg-open', folder_path])  # Linux

# Function to initialize the WebDriver based on user choice
def init_driver(browser_choice):
    try:
        if browser_choice.lower() == 'chrome':
            return webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
        elif browser_choice.lower() == 'firefox':
            return webdriver.Firefox(service=FirefoxService(GeckoDriverManager().install()))
        elif browser_choice.lower() == 'edge':
            return webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()))
    except Exception as e:
        messagebox.showerror("Error", f"Failed to open {browser_choice}. Is it installed?")
        return None

    
def on_main_window_close():
    global is_browser_open, driver

    if is_browser_open or driver:
        if messagebox.askokcancel("Quit", "Close the browser and exit?"):
            try:
                driver.quit()
            except Exception as e:
                print(f"Error closing browser: {e}")
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

# Create a queue for inter-thread communication
url_queue = queue.Queue()

def show_url_prompt():
    url_prompt_message = (
        "Enter the URL of the USIS grade uploading page:\n\n"
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
        messagebox.showerror("Error", "The entered URL is not a valid USIS mark entry URL.")
        return show_url_prompt()  # Call the function again for a valid URL

def valid_url(url):
    expected_url_part = "usis.bracu.ac.bd/academia/dashBoard/show#/academia/studentExamResult/studentExamMarksEntry"
    return url.startswith("https://") and expected_url_part in url

# def valid_url(url):
#     expected_url_part = "usis.bracu.ac.bd/academia/dashBoard/show#/academia/studentExamResult/studentExamMarksEntry"
#     return url.startswith("https://") and expected_url_part in url

def start_processing_thread():
    processing_thread = threading.Thread(target=process_gradesheet)
    processing_thread.start()

def update_text_messages(message, tag=None):
    global text_messages
    text_messages.insert(tk.END, message + "\n", tag)
    text_messages.see(tk.END)  # Auto-scroll to the end

def clear_file_path():
    entry_file_path.delete(0, tk.END)

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        # Check if the file extension is for Excel files
        if file_path.lower().endswith(('.xlsx', '.xls')):
            entry_file_path.delete(0, tk.END)
            entry_file_path.insert(0, file_path)
            label_status.config(text="", foreground="black")
            
            # Clear the text messages field
            if text_messages is not None:
                text_messages.delete('1.0', tk.END)
        else:
            messagebox.showerror("Error", "The selected file is not an Excel file.")

def enable_gui_elements():
    button_browse.config(state='normal')
    button_clear.config(state='normal')
    button_process.config(state='normal', text='Start Processing')
    entry_file_path.config(state='normal')
    browser_dropdown.config(state='readonly')

def disable_gui_elements():
    button_browse.config(state='disabled')
    button_clear.config(state='disabled')
    button_process.config(state='disabled', text='Processing started, Please wait...')
    entry_file_path.config(state='disabled')
    browser_dropdown.config(state='disabled')

def process_gradesheet():
    global driver, is_browser_open, gradesheet_url, text_messages, continue_thread_execution

    # Get values from GUI elements
    excel_path = entry_file_path.get()
    print(f"Selected File: {excel_path}")  # Debug print
    browser_choice = browser_option.get()
    print(f"Selected Browser: {browser_choice}")  # Debug print

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
    login_url = 'https://usis.bracu.ac.bd/academia/'  # Replace with your portal's login URL
    driver.get(login_url)

    # Manual login message
    messagebox.showinfo("Login", "Please log in to the portal, Navigate to the Mark Entry page. Then click OK to continue.")

    # Now schedule the URL prompt to show after login confirmation
    root.after(100, show_url_prompt)

    # Wait for the URL from the prompt
    gradesheet_url = url_queue.get()
    if gradesheet_url is None:
        # driver.quit()
        enable_gui_elements()
        return
    elif gradesheet_url:
        driver.get(gradesheet_url)
    else:
        messagebox.showerror("Error", "No URL entered. Exiting.")
        driver.quit()
        return
    
    # Check if text_messages already exists
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
    else:
        # Clear the existing text_messages widget
        text_messages.delete('1.0', tk.END)

    unmatched_students = [] # List to store unmatched students

    # Process each student's data
    for index, excel_row in df.iterrows():
        excel_student_id = str(excel_row['ID #']).strip()  # Get the student ID from Excel
        excel_grade = "{:.2f}".format(round(excel_row['Total'], 2))  # Round and format the grade
        excel_sl = str(excel_row['Sl #']).strip()  # Get the serial number
        excel_name = str(excel_row['Name']).strip()  # Get the student name

        found = False
        for portal_row in range(2, len(df) + 2):  # Iterate over rows in the portal
            xpath_id = f"/html/body/div[2]/div[1]/div[2]/div/div[3]/div[3]/div/table/tbody/tr[{portal_row}]/td[4]"
            
            try:
                portal_student_id_element = driver.find_element(By.XPATH, xpath_id)
                portal_student_id = portal_student_id_element.text.strip()

                if excel_student_id == portal_student_id:
                    update_text_messages(f"Match found for Student ID: {excel_student_id}")
                    xpath_grade_cell = f"/html/body/div[2]/div[1]/div[2]/div/div[3]/div[3]/div/table/tbody/tr[{portal_row}]/td[7]"
                    grade_cell_element = driver.find_element(By.XPATH, xpath_grade_cell)
                    driver.execute_script("arguments[0].scrollIntoView(true);", grade_cell_element)
                    grade_cell_element.click()
                    grade_input_element = grade_cell_element.find_element(By.TAG_NAME, 'input')
                    grade_input_element.clear()
                    grade_input_element.send_keys(excel_grade)
                    grade_input_element.send_keys(Keys.RETURN)
                    
                    found = True
                    break
            except NoSuchElementException:
                continue

        if not found:
            unmatched_students.append({'sl': excel_sl, 'id': excel_student_id, 'name': excel_name})
            update_text_messages(f"Student ID {excel_student_id} not found on the portal.", "unmatched")

    # # Check for unmatched students and create a log file if necessary
    # if len(unmatched_students) > 0:
    #     create_log_file(excel_base_name, unmatched_students)

    # After processing students, create a log file if there are any unmatched students
    if unmatched_students:
        log_file_path = create_log_file(excel_base_name, unmatched_students)
        user_choice = messagebox.askyesnocancel(
            "Unmatched Students Found",
            f"{len(unmatched_students)} student(s) not found. A log file has been created.\n"
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
root.title("USISphere GUI")

# Set the window icon
icon_path = 'icon.png'  # Replace with the path to your downloaded icon
root.iconphoto(False, tk.PhotoImage(file=icon_path))

# Set the size of the main window (width x height)
root.geometry("600x400")  # Adjust the size as needed

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

# Process button
button_process = tk.Button(root, text="Start Processing", command=start_processing_thread)
button_process.pack(pady=(2, 2))  # Add vertical padding below the button

# Status label
label_status = tk.Label(root, text="")
label_status.pack(pady=(2, 2))  # Add vertical padding around the label

# Bind the close event
root.protocol("WM_DELETE_WINDOW", on_main_window_close)

# Start the main loop
root.mainloop()
