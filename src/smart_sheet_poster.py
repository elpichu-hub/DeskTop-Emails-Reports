from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from time import sleep
import pandas as pd
import os
from datetime import date
from datetime import datetime
import logging
from selenium.webdriver.common.by import By

# Set up logging
log_filename = "smart_sheet_poster.log"
logging.basicConfig(filename=log_filename, level=logging.INFO, format='%(asctime)s %(levelname)s: %(message)s')

def get_time_of_day():
    current_time = datetime.now()
    hour = current_time.hour

    if 5 <= hour < 12:
        return "Morning"
    elif 12 <= hour < 17:
        return "Afternoon"
    else:
        return "Evening"

def smart_sheet_poster_function(file_path):
    # Specify the path to the chromedriver.exe executable (modify as needed)
    chrome_driver_path = r'C:\MyStuff\Projects\Developing\scraper\chromedriver-win64\chromedriver.exe'

    # Set the path to the chromedriver executable as an environment variable
    os.environ["webdriver.chrome.driver"] = chrome_driver_path

    # Initialize the WebDriver without specifying executable_path
    driver = webdriver.Chrome()

    # Read your CSV file into a Dat aFrame
    df = pd.read_csv(file_path)

    email = 'lazaro.gonzalez@conduent.com'  
    # Loop through the rows of the DataFrame
    for index, row in df.iterrows():
        sr = row['sr']
        name = row['name']
        address = row['address']
        address_2 = row['address_2']
        city = row['city']
        state = row['state']
        zip_code = row['zip_code']
        plate = row['plate']
        date_1 = row['date_1']
        date_2 = row['date_2']

        try:
            # Navigate to the URL
            driver.get("https://app.smartsheet.com/b/form/e97fb8732a8d4d7a9bae452491741c2f")
        except Exception as e:
            logging.error("Error navigating to URL", exc_info=True)
            print("Error navigating to URL", exc_info=True)
            return

        # Find and fill the form fields based on the data from the current row
        search_box_requesting_agent = driver.find_element("id", "text_box_Requesting Agent")
        search_box_requesting_agent.send_keys('Lazaro Gonzalez')
        search_box_requesting_agent.send_keys(Keys.ENTER)

        search_box_agent_email = driver.find_element("id", "text_box_Agent Email Address")
        search_box_agent_email.send_keys(email)
        search_box_agent_email.send_keys(Keys.ENTER)

        search_box_sr = driver.find_element("id", "text_box_SR Number")
        search_box_sr.send_keys(sr)
        search_box_sr.send_keys(Keys.ENTER)

        search_box_date = driver.find_element("id", "date_Date Submitted")
        today = date.today().strftime("%m/%d/%Y")
        search_box_date.send_keys(today)
        search_box_date.send_keys(Keys.ENTER)

        search_box_request_type = driver.find_element("id", "select_input_Request Type")
        search_box_request_type.send_keys("ROV Lookup")
        search_box_request_type.send_keys(Keys.ENTER)

        radio_FTE_request = driver.find_element("name", "1R1ZL7Lz")
        radio_FTE_request.click()

        send_copy = driver.find_element("name", "EMAIL_RECEIPT_CHECKBOX")
        send_copy.click()

        send_email_to = driver.find_element("id", "text_box_EMAIL_RECEIPT")
        send_email_to.send_keys(email)
        send_email_to.send_keys(Keys.ENTER)

        send_email_to = driver.find_element("id", "textarea_Agent Comments")
        message = f"""Good {get_time_of_day()},\n\nCan you please confirm if the ROV of License Plate {plate} matches the below information for transactions from ({date_1} - {date_2})?\n\n{name}\n{address}\n{address_2}\n{city}, {state} {zip_code}\n\n**If the information does not match, can you provide the correct ROV so a transfer of responsibility can be processed?**"""
        send_email_to.send_keys(message)
        send_email_to.send_keys(Keys.ENTER)

        sleep(3)  # Adjust the sleep duration if needed

        # Click the Submit button
        # Assuming 'driver' is your WebDriver instance
        submit_button = driver.find_element("css selector", "button[data-client-id='form_submit_btn']")
        submit_button.click()

        sleep(3)



    # Close the WebDriver after processing all rows
    driver.quit()
