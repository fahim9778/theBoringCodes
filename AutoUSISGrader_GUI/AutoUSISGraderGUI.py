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
driver = None  # Declare driver as a global variable


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
    global driver, is_browser_open
    is_browser_open = True

    if browser_choice.lower() == 'chrome':
        return webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
    elif browser_choice.lower() == 'firefox':
        return webdriver.Firefox(service=FirefoxService(GeckoDriverManager().install()))
    elif browser_choice.lower() == 'edge':
        return webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()))
    else:
        raise ValueError("Unsupported browser. Choose 'Chrome', 'Firefox', or 'Edge'.")
    
def on_main_window_close():
    global is_browser_open, driver
    if is_browser_open:
        if messagebox.askokcancel("Quit", "The browser maybe still open. Close it and exit?"):
            try:
                driver.quit()
            except Exception as e:
                print(f"Error closing browser: {e}")
            is_browser_open = False
    root.destroy()

# Create a queue for inter-thread communication
url_queue = queue.Queue()

def show_url_prompt():
    url_prompt_message = (
    "Enter the URL of the USIS grade uploading page:\n\n"
    "IMPORTANT: Please DO NOT interact (click, resize, move, etc.) "
    "with the browser window while processing is ongoing."
)
    url = simpledialog.askstring("Input", url_prompt_message)
    url_queue.put(url)  # Put the URL into the queue

def start_processing_thread():
    processing_thread = threading.Thread(target=process_gradesheet)
    processing_thread.start()

def update_text_messages(message, tag=None):
    global text_messages
    text_messages.insert(tk.END, message + "\n", tag)
    text_messages.see(tk.END)  # Auto-scroll to the end

def clear_file_path():
    entry_file_path.delete(0, tk.END)

def process_gradesheet():
    global driver, is_browser_open, gradesheet_url

    # Get values from GUI elements
    excel_path = entry_file_path.get()
    browser_choice = browser_option.get()

    if not excel_path or not browser_choice:
        messagebox.showerror("Error", "Please specify the gradesheet file and browser choice.")
        return

    # Check if the file exists
    if not os.path.exists(excel_path):
        messagebox.showerror("Error", "The specified file does not exist. Please check the path and try again.")
        return

    # Read the Excel file into a Pandas DataFrame
    df = pd.read_excel(excel_path)

    # Extract the base name of the Excel file (without path and extension)
    excel_base_name = os.path.splitext(os.path.basename(excel_path))[0]

    # Initialize WebDriver
    driver = init_driver(browser_choice)

    # Open the login page
    login_url = 'https://usis.bracu.ac.bd/academia/'  # Replace with your portal's login URL
    driver.get(login_url)

    # Manual login message
    messagebox.showinfo("Login", "Please log in to the portal, Navigate to the Mark Entry page. Then click OK to continue.")

    # Now schedule the URL prompt to show after login confirmation
    root.after(100, show_url_prompt)

    # Wait for the URL from the prompt
    gradesheet_url = url_queue.get()  # Retrieve the URL from the queue
    if gradesheet_url:
        driver.get(gradesheet_url)
    else:
        messagebox.showerror("Error", "No URL entered. Exiting.")
        driver.quit()
        return
    
    # Add a Text widget for displaying messages
    label_messages = tk.Label(root, text="Processing Messages:")
    label_messages.pack()

    global text_messages
    text_messages = tk.Text(root, height=10, width=50, font=("Helvetica", 10))

    # Add a scrollbar to the Text widget
    scrollbar = tk.Scrollbar(root, command=text_messages.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    text_messages.config(yscrollcommand=scrollbar.set)
    text_messages.pack()

    # Configure a tag for red and bold text
    text_messages.tag_configure("unmatched", foreground="red", font=("Helvetica", 10, "bold"))

    unmatched_students = []

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

# GUI setup
root = tk.Tk()
root.title("USISphere GUI")

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
button_browse = tk.Button(button_frame, text="Browse", command=lambda: entry_file_path.insert(0, filedialog.askopenfilename()))
button_browse.pack(side=tk.LEFT, padx=5)

# Clear Field button
button_clear = tk.Button(button_frame, text="Clear Field", command=clear_file_path)
button_clear.pack(side=tk.LEFT, padx=5)

# Browser choice dropdown
label_browser_choice = tk.Label(root, text="Select Browser:")
label_browser_choice.pack(pady=(2, 0))  # Add vertical padding above the label
browser_option = tk.StringVar(root)
browser_option.set("Firefox")  # Default value
browser_dropdown = ttk.Combobox(root, state="readonly", values=("Chrome", "Firefox", "Edge"), width=20)
browser_dropdown.current(1)  # Set default selection to Firefox
browser_dropdown.pack(pady=2)  # Add vertical padding around the dropdown

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
