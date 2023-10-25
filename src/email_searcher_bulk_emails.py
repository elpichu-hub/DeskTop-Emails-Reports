import email
import imaplib
import re
import logging
import traceback
import quopri
from datetime import datetime
import email_config
import send_email_function

# Set up logging
log_filename = "bulk_generator.log"
logging.basicConfig(filename=log_filename, level=logging.INFO, format='%(asctime)s %(levelname)s: %(message)s')

# Set up your email and password
username = email_config.EMAIL_ADDRESS_AUTO
password = email_config.EMAIL_PASSWORD_AUTO

# Set the mailbox you want to check
inbox_mailbox = "inbox"
processed_mailbox = "processed"

# Set the keyword you want to search for in the subject
keyword = 'Bulk'

# Set the mail server and port
mail_server = "imap.gmail.com"
mail_port = 993

# Connect to the server
mail = imaplib.IMAP4_SSL(mail_server, mail_port)

# Login to your account
mail.login(username, password)

# Select the mailbox you want to check
mail.select(inbox_mailbox)

# Search the mailbox for unread emails with the keyword in the subject
result, data = mail.uid('search', None, '(UNSEEN SUBJECT "{}")'.format(keyword))
if result == 'OK':
    email_ids = data[0].split()
    if not email_ids:
        logging.info("No unread emails found.")
        print("No unread emails found.")
        exit(0)
else:
    logging.error("Failed to search for unread emails.")
    print("Failed to search for unread emails.")

def create_bulk_email_string(email_body):
    # Extract all potential email addresses using a regular expression
    email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+'
    potential_emails = re.findall(email_pattern, email_body)

    # Validate and correct the emails
    corrected_email_list = []
    for email in potential_emails:
        email = email.strip()
        # If the email doesn't have a dot (.) after the @ symbol, append .com
        if '.' not in email.split('@')[1]:
            email += '.com'
        # Use a regex pattern to validate the corrected email
        match = re.match(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$', email)
        if match:
            corrected_email_list.append(email)

    # Add the three extra email addresses
    extra_emails = ['lazaro.gonzalez@conduent.com', 'lashanda.hicks@conduent.com', 'mary.rodriguez@conduent.com']
    corrected_email_list.extend(extra_emails)
    
    # Remove duplicates
    corrected_email_list = list(set(corrected_email_list))

    bcc_string = '; '.join(corrected_email_list)
    return bcc_string


# For each email that matches, get the full email data
for email_id in email_ids:
    # Fetch the full email data
    result, message_data = mail.uid('fetch', email_id, '(BODY.PEEK[])')
    if result == 'OK':
        # Check if message_data[0][1] is not None
        if message_data[0] is not None and message_data[0][1] is not None:
            raw_email = message_data[0][1].decode("utf-8")
            email_message = email.message_from_string(raw_email)

            # Get the email body
            email_body = ""
            if email_message.is_multipart():
                for part in email_message.walk():
                    if part.get_content_type() == "text/plain":
                        email_body = part.get_payload(decode=True)  # decode
            else:
                email_body = email_message.get_payload(decode=True)  # decode

            # Decode the body from quoted-printable to utf-8
            email_body = quopri.decodestring(email_body).decode('utf-8')

            # Remove mailto links
            email_body = re.sub(r'<mailto:[^>]*>', '', email_body)

            # Remove the signature
            signature_index = email_body.find("Thank You,")
            if signature_index != -1:
                email_body = email_body[:signature_index]

            # print("Email body: ", email_body)
            logging.info("Email body: %s", email_body)

            bulk_addresses = create_bulk_email_string(email_body)
            print(bulk_addresses)

            # Move the email to the "processed" mailbox
            result = mail.uid('COPY', email_id, processed_mailbox)
            if result[0] == 'OK':
                # Mark the original email for deletion
                mail.uid('STORE', email_id, '+FLAGS', '\\Deleted')
                mail.expunge()

            signature = email_config.team_lead_elite_signature
            subject = 'SunPass Survey Acknowledgement'
            body = f"""<h3>Hi, below you can find the emails for the surveys acknowledged:</h3>
            <p>{bulk_addresses}</p>
            {signature}
            """

            if bulk_addresses:
                try:
                    send_email_function.send_email(f'Bulk Email String Success', 'lazaro.gonzalez@conduent.com', f'Bulk Email String {datetime.now()}')
                    send_email_function.send_email(f'Bulk Email String {datetime.now()}', 'lazaro.gonzalez@conduent.com', f'{bulk_addresses}')
                    send_email_function.send_email(subject, 'lazaro.gonzalez@conduent.com', body)
                    logging.info(f'Bulk Email String, succesful')
                    print('this ran')
                except Exception as e:
                    error_message = str(e)  # Get the error message as a string
                    traceback_str = traceback.format_exc()  # Get the traceback as a string
                    error_message_with_traceback = f"{error_message}\n\nTraceback:\n{traceback_str}"
                    send_email_function.send_email(f'Error SBS', 'lazaro.gonzalez@conduent.com', error_message_with_traceback)
                    logging.error('Failed to send SBS email.', exc_info=True)
            else:
                send_email_function.send_email(f'no Bulk Emails', 'lazaro.gonzalez@conduent.com', f'Bulk Email String {datetime.now()}')
            

# Close the mailbox and log out
mail.close()
mail.logout()

# Log script completion
logging.info("Script execution completed.")
print("Script execution completed.")
