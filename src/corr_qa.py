import openpyxl
import pyautogui
import time
import os
import re
import email_config
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders
import encouraging_and_success_messages
import random
from datetime import datetime
import pygetwindow as gw


def send_email(subject, recipient, body, img_path1=None, img_path2=None, corr_dispute_form=None):
    
    EMAIL_ADDRESS = email_config.EMAIL_ADDRESS_AUTO
    EMAIL_PASSWORD = email_config.EMAIL_PASSWORD_AUTO

    # Create the email message
    message = MIMEMultipart()
    message['From'] = EMAIL_ADDRESS
    message['To'] = recipient
    message['Subject'] = subject

    message.attach(MIMEText(body, 'html'))

    # If an image path is provided, add the image as an inline attachment
    if img_path1 is not None:
        with open(img_path1, 'rb') as img_file:
            img_data = img_file.read()
        img_mime = MIMEImage(img_data)
        img_mime.add_header('Content-ID', '<{}>'.format(os.path.basename(img_path1)))
        img_mime.add_header('Content-Disposition', 'inline', filename=os.path.basename(img_path1))
        message.attach(img_mime)

    # If an img_path_100 is provided, add the image as an inline attachment
    if img_path2 is not None:
        with open(img_path2, 'rb') as img_file:
            img_data = img_file.read()
        img_mime = MIMEImage(img_data)
        img_mime.add_header('Content-ID', '<{}>'.format(os.path.basename(img_path2)))
        img_mime.add_header('Content-Disposition', 'inline', filename=os.path.basename(img_path2))
        message.attach(img_mime)

    if corr_dispute_form is not None:
        with open(corr_dispute_form, 'rb') as mentors_file:
            mentors_part = MIMEBase('application', 'octet-stream')
            mentors_part.set_payload(mentors_file.read())
            encoders.encode_base64(mentors_part)
            mentors_part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(corr_dispute_form))
            message.attach(mentors_part)

    # Connect to the Gmail SMTP server and send the email
    with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
        smtp.starttls()
        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        smtp.send_message(message)


# Replace 'your_file.xlsx' with the path to your Excel file
# file_path = os.path.abspath('data/Correspondence Grading Form B.xlsx')


def sort_correspondence_qa(file_path, my_correspondence_agents=None):
    try:
        # Load the workbook with data_only=True to retrieve cell values
        workbook = openpyxl.load_workbook(file_path, data_only=True)

        os.startfile(file_path)

        # Wait for Excel to open
        time.sleep(5)

        # Find the window
        excel_windows = gw.getWindowsWithTitle('Correspondence Grading Form B')  # Adjust title for your specific Excel file
        if excel_windows:
            excel_window = excel_windows[0]  # Assuming the first window is the one you want
            excel_window.activate()
            pyautogui.moveTo(excel_window.left + 300, excel_window.top + 300)

        for sheet in workbook:
            agent_name = sheet['E2'].value
            cleaned_name = re.sub(r'[^A-Za-z0-9]', '', agent_name)

            # Delay to ensure Excel is active window
            time.sleep(2)

            # Take a screenshot of the initial view
            screenshot = pyautogui.screenshot()
            # Specify the region for the first screenshot
            excel_region = (20, 225, 1490, 975)  # Adjust as necessary
            # Crop and save the first screenshot
            first_screenshot = screenshot.crop(excel_region)
            first_screenshot_path = f'data/screenshots/first_excel_screenshot_{cleaned_name}.png'
            first_screenshot.save(first_screenshot_path)
            print(f"Screenshot saved as 'data/screenshots/first_excel_screenshot:{cleaned_name}.png'")

            # Scroll down
            pyautogui.scroll(-1000)
            # Wait a moment for the scroll to complete
            time.sleep(1)

            # Take a screenshot of the new view after scrolling
            screenshot_after_scroll = pyautogui.screenshot()
            # Specify the region for the second screenshot
            excel_region_after_scroll = (20, 540, 1490, 975)  # Adjust as necessary
            # Crop and save the second screenshot
            second_screenshot = screenshot_after_scroll.crop(excel_region_after_scroll)
            second_screenshot_path = f'data/screenshots/second_excel_screenshot_{cleaned_name}.png'
            second_screenshot.save(second_screenshot_path)
            print(f"Screenshot saved as 'data/screenshots/second_excel_screenshot{cleaned_name}.png'")

            # Switch to the next sheet using Ctrl+PageDown
            pyautogui.keyDown('ctrl')
            pyautogui.press('pagedown')
            pyautogui.keyUp('ctrl')
            # Wait a moment for the sheet switch to complete
            time.sleep(1)

            # Get the current date and time
            now = datetime.now()
            greeting = ""

            # Determine the appropriate greeting based on the current time
            if now.hour < 12:
                greeting = "Good morning"
            elif now.hour < 18:
                greeting = "Good afternoon"
            else:
                greeting = "Good evening"

            subject = f"Correspondence QA for {agent_name}"
            recipient = 'lazaro.gonzalez@conduent.com'
            corr_dispute_form = "data/DISPUTE FORM CORRESPONDENCE.docx"
        
           
            date_range = sheet['J2'].value
            result_string = ""

            # Get the average score from the Excel sheet
            average_score_cell_value = sheet['G4'].value
            print(f"Average score: {average_score_cell_value}")

            # Initialize result_string
            result_string = ""

            # Convert average score to a number and remove the percentage sign if present
            try:
                # If the value is a string with a percentage, strip the '%' and convert
                if isinstance(average_score_cell_value, str):
                    average_score = float(average_score_cell_value.strip('%') * 100)
                else:
                    average_score = float(average_score_cell_value * 100)
            except ValueError:
                print("Average score is not a number.")
            else:
                # Use the average_score as a number hereafter
                if average_score >= 90:
                    random_success_message = random.choice(encouraging_and_success_messages.success_messages)
                    result_string = f"<h3>{greeting} {agent_name},\n\nYour average QA for the week of {date_range} is <span style='color: green;'>{average_score}%</span>. {random_success_message}\n</h3>"
                else:
                    random_encouraging_message = random.choice(encouraging_and_success_messages.encouraging_messages)
                    result_string = f"<h3>{greeting} {agent_name},\n\nYour average QA for the week of {date_range} is <span style='color: red;'>{average_score}%</span>. <span style='background-color: #59FFA0;'>{random_encouraging_message}</span>\n</h3>"

            print(result_string)
            
            body = f"""<html>
                <body>
                {result_string}<br>
                <img src='cid:{os.path.basename(first_screenshot_path)}' alt='Image 1' style='width: 1000px; height: 600px;'><br>
                <img src='cid:{os.path.basename(second_screenshot_path)}' alt='Image 2' style='width: 1000px; height: 300px;'><br>
                </body>
            </html>"""

            if my_correspondence_agents is not None:
                print('my_correspondence_agents is not None')
                if agent_name in my_correspondence_agents:
                    send_email(subject, recipient, body, first_screenshot_path, second_screenshot_path, corr_dispute_form)
                    print(f"Email sent to {recipient}.")
            else:
                print('my_correspondence_agents is None')
                send_email(subject, recipient, body, first_screenshot_path, second_screenshot_path, corr_dispute_form)
                print(f"Email sent to {recipient}. no agents names provided")

        time.sleep(5)

        # Close Excel (Windows-specific)
        if os.name == 'nt':
            os.system('taskkill /F /IM excel.exe')
            print(f"Excel file '{workbook}' closed.")
        else:
            print("This method is only supported on Windows.")

    except FileNotFoundError:
        print(f"File '{file_path}' not found.")
    except Exception as e:
        print(f"An error occurred: {e}")



# sort_correspondence_qa(file_path, email_config.my_correspondence_agents)
# sort_correspondence_qa(file_path)