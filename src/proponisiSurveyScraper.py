from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from time import sleep
import os
from datetime import datetime, timedelta
import email_config

# This function sends an email with an attachment
def send_email(subject, recipient, body, file_path=None):
    import smtplib
    from email.mime.text import MIMEText
    from email.mime.multipart import MIMEMultipart
    import os
    from email.mime.base import MIMEBase
    from email import encoders
    

    EMAIL_ADDRESS = email_config.EMAIL_ADDRESS_AUTO
    EMAIL_PASSWORD = email_config.EMAIL_PASSWORD_AUTO

    # Create the email message
    message = MIMEMultipart()
    message['From'] = EMAIL_ADDRESS
    message['To'] = recipient
    message['Subject'] = subject

    message.attach(MIMEText(body, 'html'))

    if file_path is not None:
        with open(file_path, 'rb') as file:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(file.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(file_path))
            message.attach(part)
    

    # Connect to the Gmail SMTP server and send the email
    with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
        smtp.starttls()
        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        smtp.send_message(message)


# This function returns the formatted dates for the query
def get_dates_for_query():
    # Get today's date and time
    today = datetime.today()
    two_days_ago = today - timedelta(days=2)
    two_days_ago = two_days_ago.replace(hour=23, minute=59, second=00, microsecond=000000)
    one_months_ago = today - timedelta(days=30)
    one_months_ago = one_months_ago.replace(hour=00, minute=00, second=00, microsecond=000000)
    # Format the date and time
    formatted_date_two_days_ago = two_days_ago.strftime("%m/%d/%Y %I:%M %p")
    formatted_date_one_month_ago = one_months_ago.strftime("%m/%d/%Y %I:%M %p")
    print('Getting dates for query... Worked!')
    if today.weekday() == 0:
        four_days_ago = today - timedelta(days=4)
        four_days_ago = four_days_ago.replace(hour=23, minute=59, second=00, microsecond=000000)
        formatted_date_four_days_ago = four_days_ago.strftime("%m/%d/%Y %I:%M %p")
        print(formatted_date_four_days_ago)
        return formatted_date_four_days_ago, formatted_date_one_month_ago
    return formatted_date_two_days_ago, formatted_date_one_month_ago

# This function navigates to the login page, logs in, sets the dates and downloads the file
def navigate_login_download(formatted_date_two_days_ago, formatted_date_one_month_ago):

    download_dir = r'C:\MyStuff\Projects\Developing\Tests\StatsOrginizer\src\downloads_proponisi'
    chrome_options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)
    # Specify the path to the chromedriver.exe executable (modify as needed)
    chrome_driver_path = r'C:\MyStuff\Projects\Developing\scraper\chromedriver-win64\chromedriver.exe'

    # Set the path to the chromedriver executable as an environment variable
    os.environ["webdriver.chrome.driver"] = chrome_driver_path

    # Initialize the WebDriver without specifying executable_path
    driver = webdriver.Chrome(options=chrome_options)

    # Navigate to the URL
    driver.get("https://www.tpcdm.com/Account/Login?ReturnUrl=%2f")

    # Find and fill the form fields based on the data from the current row
    username_field = driver.find_element("id", "Email")
    username_field.send_keys('lazaro.gonzalez@conduent.com')
    username_field.send_keys(Keys.ENTER)

    password_field = driver.find_element("id", "Password")
    password_field.send_keys('Pinkfloyd(07)!')
    password_field.send_keys(Keys.ENTER)

    driver.get("https://www.tpcdm.com/Ticketing/Report/TicketExport")

    print('Logged In!')

    sleep(5)  # Adjust the sleep duration if needed

    start_date_field = driver.find_element("id", "StartDate")
    start_date_field.clear()
    start_date_field.send_keys(formatted_date_one_month_ago)
    start_date_field.send_keys(Keys.ENTER)
    print(formatted_date_one_month_ago)

    # Execute JavaScript to blur the element and remove focus
    driver.execute_script("arguments[0].blur();", start_date_field)

    end_date_field = driver.find_element("id", "EndDate")
    end_date_field.clear()
    end_date_field.send_keys(formatted_date_two_days_ago)
    start_date_field.send_keys(Keys.ENTER)
    print(formatted_date_two_days_ago)

    # Execute JavaScript to blur the element and remove focus
    driver.execute_script("arguments[0].blur();", end_date_field)

    print('Dates set!')

    # Click the Submit button
    submit_button = driver.find_element("id", "btnSubmit")
    submit_button.click()

    print('Submitted! Wait 30 seconds...')

    sleep(30)  # Adjust the sleep duration if needed

    print('Waited 30 seconds!')

    # Close the WebDriver after processing all rows
    driver.quit()

    print('Closed WebDriver!')

    # Rename the downloaded file
    downloaded_file_path = os.path.join(download_dir, 'TicketExport.csv')
    date_for_name = datetime.now().strftime("%Y%m%d%H%M%S")
    os.rename(downloaded_file_path, os.path.join(download_dir, 'TicketExport_' + date_for_name + '.csv'))
    csv_file_path = download_dir + '\\' + 'TicketExport_' + date_for_name + '.csv'

    # Return the path to the downloaded file and the download directory
    return csv_file_path, download_dir

# This function cleans the data
def clean_data(csv_file_path, download_dir):
    import pandas as pd

    # Load the data from the CSV file
    df = pd.read_csv(csv_file_path, index_col=False)

    # Keep only the specified columns
    df = df[['Ticket Status', 'Ticket Resolution', 'Agent', 'Agent', 'Supervisor', 'Survey', 'Preview Link', 'Created Date', ' Interaction Id']]

    # filter rows where Ticket Status is not 'Close'
    not_closed_tickets_df = df[df['Ticket Status'] != 'Close']

    # Export the data to a CSV file and rename it
    file_name = f'not_closed_tickets${datetime.now().strftime("%Y%m%d%H%M%S")}.csv'
    file_path = os.path.join(download_dir, file_name)
    not_closed_tickets_df.to_csv(file_path, index=False, encoding='utf-8-sig', index_label=False)
    print('Tickets not closed exported!: ')
    print(not_closed_tickets_df)

    if not_closed_tickets_df.empty:
        print('No tickets not closed!')
        return None

    return file_path

# This function sends an email with an attachment
def main_func():
    # Check if today is Sunday
    if datetime.today().weekday() == 6 or datetime.today().weekday() == 5:
        print(f'Today is {datetime.today().weekday()}, so no email will be sent')
        send_email(f'Today is {datetime.today().weekday()}, No Surveys', 'lazaro.gonzalez@conduent.com', f'Today is {datetime.today().weekday()}, No Surveys', None)
        return
    else:
        try:
            # Get the formatted dates for the query
            two_or_four_days_ago, formatted_date_one_month_ago = get_dates_for_query()
            # Navigate to the login page, log in, set the dates and download the file
            csv_file_path, download_dir = navigate_login_download(two_or_four_days_ago, formatted_date_one_month_ago)
            # Clean the data
            file_path = clean_data(csv_file_path, download_dir)
            # Send the email if there are tickets not closed else send email with no attachment
            # explain all tickets are closed
            if file_path is None:
                send_email(f'No tickets not closed', 'lazaro.gonzalez@conduent.com', f'All Tickets Due Are Closed {datetime.now()}')
            else:
                send_email(f'Tickets not closed', 'lazaro.gonzalez@conduent.com', 'Tickets not closed', file_path)
        # If there is an error, send an email with the error message
        except Exception as e:
            print(str(e))
            send_email(f'Error for Tickets not closed', 'lazaro.gonzalez@conduent.com', str(e), None)
        

main_func()



    

