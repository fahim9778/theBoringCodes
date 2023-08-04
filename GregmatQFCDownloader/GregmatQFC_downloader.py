# Project: Gregmat Quant Flashcard Downloader
# Author: MD. Fakhruddin Gazzali Fahim
# Date created: 04-Aug-2023 3:34 PM

"""
This script downloads all the flashcards from the Gregmat Quant Flashcards course.
The script assumes that the user has Gregmat+ subscription and is logged in.
The script also assumes that the user has Firefox installed,
since the script uses the Firefox webdriver.
"""


import os
import time
import urllib.request
from selenium import webdriver
from selenium.webdriver.common.by import By
# from selenium.webdriver.common.keys import Keys
# from webdriver_manager.firefox import GeckoDriverManager


# Determine the directory where the .py file is located
current_directory = os.path.dirname(os.path.realpath(__file__))

# Create the parent folder if it doesn't exist
parent_folder = os.path.join(current_directory, "Gregmat Quant Flashcards")
if not os.path.exists(parent_folder):
    os.makedirs(parent_folder)

driver = webdriver.Firefox()

# Currently, there are 17 pages or groups of flashcards
current_flashcard_groups = 17 # Change this as needed

# Loop over the pages
for i in range(1, current_flashcard_groups+1):  # Adjust the range as needed
    # Navigate to the page
    driver.get(f"https://www.gregmat.com/course/quant-flashcards-group-{i}")

    # Wait for a bit for the page to load
    time.sleep(5) # Adjust this as needed

    # Create a subfolder for the images from this page
    folder = os.path.join(parent_folder, f'Flashcard_Group_{i}')
    if not os.path.exists(folder):
        os.makedirs(folder)

    # Find all images
    images = driver.find_elements(By.TAG_NAME, 'img')

    # Download each image
    for j, img in enumerate(images):
        url = img.get_attribute('src')
        urllib.request.urlretrieve(url, f'{folder}/flashcard_{j}.png')

driver.quit()  # Close the browser when done
