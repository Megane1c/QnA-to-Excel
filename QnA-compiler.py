# -*- coding: utf-8 -*-
"""
Created on Sun Jul  9 13:56:48 2023

@author: Megane
"""

from PIL import Image
from io import BytesIO
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time
import base64
import os


# Instance of firefox driver
driver = webdriver.Firefox()

# Maximize the window
driver.maximize_window()

# Go to log in page, will hide which site
driver.get(website)

# Sleep is used to ensure that the webpage fully loads
time.sleep(2)

'''
For this part, one should log in their respective accounts
on the webpage, then hit enter in the console once the
page loads
'''
input('Press enter after logging in')

# Exam page, again site is hidden
driver.get(website)

time.sleep(2)

# Iterate over the table pages
while True:
    # Scroll down using JavaScript
    driver.execute_script("window.scrollBy(0, window.innerHeight)")
    
    # Wait for the webpage to fully load
    wait = WebDriverWait(driver, 10)
    wait.until(EC.presence_of_element_located((By.ID, "tbl-taken")))
    
    # Locate all rows in the table
    table_rows = driver.find_elements(By.XPATH, "//table[@id='tbl-taken']/tbody/tr")
    
    # Locate "Next" button
    next_button = driver.find_element(By.ID, "tbl-taken_next")
    
    time.sleep(2)
    
    # Loop through the items with review button
    for review_row in table_rows:
        
        # Get the name of the subject
        subject_name = review_row.find_element(By.XPATH, "./td[2]").text
        
        # Create a folder for the subject if it does not exist, skip creation if it exists
        subject_folder = os.path.join(".", subject_name)
        if os.path.exists(subject_folder):
            continue
        
        os.makedirs(subject_folder, exist_ok=True)
        
        # Scroll using JavaScript to bring the button into view
        review_button = review_row.find_element(By.XPATH, ".//a[contains(@href, 'exam/review')]")
        driver.execute_script("arguments[0].scrollIntoView();", review_button)
        
        # Current window handle
        current_window = driver.current_window_handle
        
        # Open in new tab
        review_button.send_keys(Keys.CONTROL + Keys.RETURN)
        
        # Switch to new tab
        driver.switch_to.window(driver.window_handles[-1])
        
        time.sleep(2)
        
        # Extract the table rows
        table_review_rows = driver.find_elements(By.XPATH, "//table[@id='tblreview']/tbody/tr")

        # Lists to store extracted data
        items = []
        questions = []
        answers = []
        question_image_filenames = []
        solution_image_filenames = []

        # Iterate
        for row in table_review_rows:
            
            # Extract item, question, and answer
            item = row.find_element(By.XPATH, "./td[1]").text
            question_element = row.find_element(By.XPATH, "./td[2]")
            question = question_element.text
            correct_answer = row.find_element(By.XPATH, "./td[3]").text
            
            # Set default image filenames
            question_image_filename = ""
            solution_image_filename = ""
            
            # Flag to keep track of the images of Solution and Question
            solution_image_downloaded = False
            question_image_downloaded = False
            
            try:
                # Image source of question
                question_image_src = question_element.find_element(By.XPATH, "./p/img").get_attribute("src")

                # Download question image
                if question_image_src and not question_image_downloaded:
                    # Image is in base64 string, needs to decode 
                    base64_image = question_image_src.split(",")[1]
                    image_data = base64.b64decode(base64_image)

                    # Create image from decoded data
                    image = Image.open(BytesIO(image_data))

                    # Save image
                    question_image_filename = f"image_question_{item}.png"

                    # Set image path to question images inside the subject
                    question_image = os.path.join(subject_folder, "question_images")

                    # Create if directory does not exist
                    if not os.path.exists(question_image):
                        os.makedirs(question_image)
                    
                    question_image_path = os.path.join(subject_folder, "question_images", question_image_filename)
                    
                    # Save the image
                    image.save(question_image_path)
                    
                    # Set to true
                    question_image_downloaded = True
                    
                    time.sleep(2)
                    
            except NoSuchElementException:
                pass
            
            # Set to None to handle exception
            solution_button = None
            
            try:
                # Find solution button
                solution_button = row.find_element(By.XPATH, ".//a[contains(@href, '#solution')]")
                
                # Scroll to bring it into view with JavaScript
                driver.execute_script("arguments[0].scrollIntoView();", solution_button)
                
                time.sleep(2)
                
                # Use JavaScript to click the button
                driver.execute_script("arguments[0].click();", solution_button)
                
                time.sleep(2)
                
            except NoSuchElementException:
                pass
            
            try:
                # Check if the solution contains an image
                solution_element = row.find_element(By.XPATH, ".//div[contains(@class, 'card-body')]//img")
                solution_image_src = solution_element.get_attribute("src")

                # Download the solution
                if solution_image_src and not solution_image_downloaded:
                    # Image is in base64 string, needs to decode 
                    base64_image = solution_image_src.split(",")[1]
                    image_data = base64.b64decode(base64_image)

                    # Create image from decoded data
                    image = Image.open(BytesIO(image_data))

                    # Save image
                    solution_image_filename = f"image_solution_{item}.png"
                    
                    # Set image path to solution images inside the subject
                    solution_image = os.path.join(subject_folder, "solution_images")

                    # Create if directory does not exist
                    if not os.path.exists(solution_image):
                        os.makedirs(solution_image)
                    
                    # Set the image path in the subject's solution images folder
                    solution_image_path = os.path.join(subject_folder, "solution_images", solution_image_filename)
        
                    # Save the image
                    image.save(solution_image_path)
                    
                    # Set to true
                    solution_image_downloaded = True
                    
                    # Click the button again to close the solution dropdown
                    if solution_button is not None:
                        driver.execute_script("arguments[0].click();", solution_button)
                        
                    time.sleep(2)
                    
            except NoSuchElementException:
                
                # If there is no image, just click the button again to close the solution dropdown
                if solution_button is not None:
                    driver.execute_script("arguments[0].click();", solution_button)
                    
                time.sleep(2)
                
            
            # Append to the list
            items.append(item)
            questions.append(question)
            question_image_filenames.append(question_image_filename)
            answers.append(correct_answer)
            solution_image_filenames.append(solution_image_filename)
            
        # Dataframe
        data = pd.DataFrame({
            "Item": items,
            "Question": questions,
            "Question Image": question_image_filenames,
            "Correct Answer": answers,
            "Solution": solution_image_filenames
        })
        
        # Save to Excel file
        output_filename = f"{subject_name}.xlsx"
        output_path = os.path.join(subject_folder, output_filename)
        data.to_excel(output_path, index=False, encoding='utf-8')
        
        # Close current tab
        driver.close()
        
        # Switch to main tab
        driver.switch_to.window(current_window)
        
        time.sleep(2)
        
        # Scroll down the page using JavaScript
        driver.execute_script("window.scrollBy(0, window.innerHeight)")
        
        time.sleep(2)
        
        
    # This checks the attribute of the next button, break the loop if it reaches
    # the end
    if "disabled" in next_button.get_attribute("class"):
        break  
    
    # Click next button
    next_button.click()

# Go to site's homepage
driver.get(website)

time.sleep(2)

# Logout
logout_button = driver.find_element(By.XPATH, "//a[contains(@href, '/auth/logout')]")
logout_button.click()

time.sleep(5)

# Close the browser
driver.quit()
