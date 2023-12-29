import os
import datetime

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.firefox.service import Service as FirefoxService
from selenium.webdriver.edge.service import Service as EdgeService
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.firefox import GeckoDriverManager

# Prompt user for the location of the Gradesheet file
excel_path = input("Please enter the path to the Gradesheet file: ").strip(' "\'')  # Strip spaces, single, and double quotes"G:\My Drive\Fall 23 OBEs\Finalized Gradesheets\CSE221 Section 08 Fall 23 Gradesheet - FGZ - Copy.xlsx"

# Check if the file exists
if not os.path.exists(excel_path):
    print("Error: The specified file does not exist. Please check the path and try again. Press Enter to exit.")
    exit(1)  # Exit the script if the file doesn't exist

# Read the Excel file into a Pandas DataFrame
df = pd.read_excel(excel_path)  # Assuming the Excel file is already read into df

# Extract the base name of the Excel file (without path and extension)
excel_base_name = os.path.splitext(os.path.basename(excel_path))[0]

# Create or open a log file with the current date and time and Excel file name in its name
current_time = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

# Get the directory where the script is located
script_dir = os.path.dirname(os.path.realpath(__file__))

# Create the log file path in the same directory as the script
log_file_path = os.path.join(script_dir, f"unmatched_students_log_{excel_base_name}_{current_time}.txt")

with open(log_file_path, "w") as log_file:
    log_file.write("Unmatched Student Log\n")
    log_file.write("Time: " + current_time + "\n")
    log_file.write("SL of Gradesheet file, Student ID, Name\n")


    # Function to initialize the WebDriver based on user choice
    def init_driver(browser_choice):
        if browser_choice.lower() == 'chrome':
            return webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
        elif browser_choice.lower() == 'firefox':
            return webdriver.Firefox(service=FirefoxService(GeckoDriverManager().install()))
        elif browser_choice.lower() == 'edge':
            return webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()))
        else:
            raise ValueError("Unsupported browser. Choose 'Chrome' or 'Firefox'.")

    # User selects the browser
    browser_choice = input("Select your browser (Chrome/Edge/Firefox): ")

    # Initialize WebDriver
    driver = init_driver(browser_choice)

    # Open the login page
    login_url = 'https://usis.bracu.ac.bd/academia/'  # Replace with your portal's login URL
    driver.get(login_url)

    # Wait for manual login
    input("Please log in to the portal. After logging in, press Enter to continue...")

    # Ask user for the gradesheet URL
    gradesheet_url = input("Enter the URL of the USIS grade uploading page: ")
    driver.get(gradesheet_url)


    # Iterate over each student's data in the Excel file
    for index, excel_row in df.iterrows():
        excel_student_id = str(excel_row['ID #']).strip()  # Get the student ID from Excel

        # Round the grade to two decimal points before converting to string to avoid floating point errors in the portal
        excel_grade = "{:.2f}".format(round(excel_row['Total'], 2)) # Get the grade from Excel, round to 2 decimal points and convert to string
        
        excel_sl = str(excel_row['Sl #']).strip()  # Get the serial number
        excel_name = str(excel_row['Name']).strip()  # Get the student name

        found = False
        # Adjust the range as needed
        for portal_row in range(2, len(df) + 2): # Iterate over each row of the portal table, ensures this checks all rows in the portal
            # Construct the XPath for the student ID on the portal
            xpath_id = f"/html/body/div[2]/div[1]/div[2]/div/div[3]/div[3]/div/table/tbody/tr[{portal_row}]/td[4]"
            
            try:
                portal_student_id_element = driver.find_element(By.XPATH, xpath_id)
                portal_student_id = portal_student_id_element.text.strip()

                if excel_student_id == portal_student_id:
                    print(f"Match found for Student ID: {excel_student_id}")
                    
                    # Construct the XPath for the grade cell (not the input field yet)
                    xpath_grade_cell = f"/html/body/div[2]/div[1]/div[2]/div/div[3]/div[3]/div/table/tbody/tr[{portal_row}]/td[7]"
                    grade_cell_element = driver.find_element(By.XPATH, xpath_grade_cell)
                    
                    # Click the grade cell to make it editable
                    grade_cell_element.click()
                    
                    # After the click, the input field should now be present, find it by its tag name
                    grade_input_element = grade_cell_element.find_element(By.TAG_NAME, 'input')
                    grade_input_element.clear()
                    
                    # Enter the grade and simulate pressing the Enter key
                    grade_input_element.send_keys(excel_grade)
                    grade_input_element.send_keys(Keys.RETURN)  # Simulate Enter key press
                    
                    found = True
                    break  # Exit the loop after processing the current student
            except NoSuchElementException:
                # If the student ID element is not found, it might be at the end of the table, or the ID is not present.
                print(f"Student ID element {excel_student_id} not found on the portal for row {portal_row}.")
                continue  # Check the next row

        if not found:
            print(f"Student ID {excel_student_id} not found on the portal after checking all rows.")
            # Write the unmatched student details to the log file
            log_file.write(f"{excel_student_id}, {excel_sl}, {excel_name}\n")
            print(f"(SL {excel_sl}, Student ID {excel_student_id}, Name: {excel_name}) not found on the portal and logged.")

        # # Pause to review before moving to the next entry
        # input("Press Enter to continue to the next student...")

# Keep the browser open
print("All entries processed. Review the entries in the browser.")
print(f"If any Unmatched students are found, they are logged in '{log_file_path}'.")
