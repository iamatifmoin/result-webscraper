from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from docx import Document

from selenium.webdriver.chrome.service import Service 
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options

options = Options()
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options) 

options.add_experimental_option('excludeSwitches', ['enable-logging'])

driver = webdriver.Chrome(
    service= Service(), 
    options=options,
)

start_roll = 720001
end_roll = 720077
roll_numbers = [str(i) for i in range(start_roll, end_roll + 1)]

# Create a new Word document
doc = Document()

# Iterate through each roll number
for roll_number in roll_numbers:
    driver.get('https://result.npgc.in/')  # Replace with the actual webpage URL
    # Find the input field and enter the roll number
    input_field = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, 'ctl00_ContentPlaceHolder1_txtStudentRollNo'))
    )
    input_field.clear()
    input_field.send_keys(roll_number)
    input_field.send_keys(Keys.RETURN)


    # Take a screenshot of the result
    screenshot_path = f'screenshot_{roll_number}.png'  # Generate a unique screenshot path for each roll number
    driver.save_screenshot(screenshot_path)

    # Add the screenshot to the Word document
    doc.add_picture(screenshot_path)

# Save the Word document
doc.save('result_screenshots.docx')  # Provide the desired file path and name for the Word document

# Close the browser
driver.quit()