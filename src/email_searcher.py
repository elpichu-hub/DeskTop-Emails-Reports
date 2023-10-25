import email
import imaplib
import os
import email_config
from datetime import datetime
import QA
import logging
import re
import send_email_function
import traceback

# Set up logging
log_filename = "script.log"
logging.basicConfig(filename=log_filename, level=logging.INFO, format='%(asctime)s %(levelname)s: %(message)s')


# save the attachments' path from the email to the pass it to a function
def save_attachments(email_message, sub_folder_path):
    file_paths = []
    priority_filename = "agents_list_do_not_delete.csv"
    priority_file_path = None
    

    # If the email has attachments
    for part in email_message.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue

        filename = part.get_filename()

        if bool(filename):
            # Only save .csv attachments
            if not filename.lower().endswith('.csv'):
                continue
            file_path = os.path.join(sub_folder_path, filename)

            # If this is the priority file, save its path separately
            if filename == priority_filename:
                priority_file_path = file_path
            else:
                file_paths.append(file_path)

            if not os.path.isfile(file_path):
                logging.info('Saving file: %s', filename)
                print('Saving file: %s', filename)
                try:
                    with open(file_path, 'wb') as f:
                        f.write(part.get_payload(decode=True))
                except Exception as e:
                    logging.error('Error saving file: %s', filename, exc_info=True)
                    print('Error saving file: %s', filename, exc_info=True)

    # If we found the priority file, prepend its path to the list
    if priority_file_path:
        file_paths.insert(0, priority_file_path)

    return file_paths

# Set up your email and password
username = email_config.EMAIL_ADDRESS_AUTO
password = email_config.EMAIL_PASSWORD_AUTO

# Set the mailbox you want to check
inbox_mailbox = "inbox"
processed_mailbox = "processed"

# Set the keyword you want to search for in the subject
keyword = 'Quality Assurance'

# Set the mail server and port
mail_server = "imap.gmail.com"
mail_port = 993

# Set the folder name for saving attachments
folder_name = "automated_emails"

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

# Load the processed email IDs from a file or initialize an empty set
processed_ids = set()
if os.path.isfile('processed_ids.txt'):
    with open('processed_ids.txt', 'r') as f:
        processed_ids = set(f.read().splitlines())

# Create the folder if it doesn't exist
if not os.path.exists(folder_name):
    os.makedirs(folder_name)

file_paths = None
email_subject = None
date_range = None
email_address = None
signature = None

# For each email that matches, get the full email data
for email_id in email_ids:
    if email_id not in processed_ids:
        # Fetch the full email data
        result, message_data = mail.uid('fetch', email_id, '(BODY.PEEK[])')
        if result == 'OK':
            # Check if message_data[0][1] is not None
            if message_data[0] is not None and message_data[0][1] is not None:
                raw_email = message_data[0][1].decode("utf-8")
                email_message = email.message_from_string(raw_email)

                # Get the email sender
                email_sender = email.utils.parseaddr(email_message['From'])[1]
                logging.info("Email sender: %s", email_sender)
                print("Email sender: %s", email_sender)

                # Get the email subject
                email_subject = email_message['Subject']
                logging.info("Email subject: %s", email_subject)
                print("Email subject: %s", email_subject)

                # Remove the prefix "Fwd: " if it exists
                if email_subject.startswith('Fwd: '):
                    email_subject = email_subject[5:]

                # Extract the date range from the subject
                date_range = email_subject.split('Quality Assurance ')[1].split(' ')[0:3]
                date_range = ' '.join(date_range)
                logging.info("Date Range: %s", date_range)
                print("Date Range: %s", date_range)

                # Extract the email address from the subject
                email_match = re.search(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', email_subject)
                if email_match:
                    email_address = email_match.group(0)
                else:
                    # If no email found in the subject, the email will be sent to the sender
                    email_address = email_sender
                
                logging.info("Email: %s", email_address)
                print("Email: %s", email_address)
                
                # check the email address to assign a signature
                if email_address == 'Edgar.Carneiro@conduent.com' or email_address == 'carneiroedgar13@gmail.com' or email_address == 'edgar.carneiro@conduent.com':
                    signature = email_config.team_lead_edgar_signature
                elif email_address == 'karina.morales@conduent.com':
                    signature = email_config.team_lead_karina_signature
                else:
                    signature = None

                # Create a subfolder with the current time
                sub_folder_name = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
                sub_folder_path = os.path.join(folder_name, sub_folder_name)
                if not os.path.exists(sub_folder_path):
                    os.makedirs(sub_folder_path)

                file_paths = save_attachments(email_message, sub_folder_path)

                # # Move the email to the "processed" mailbox
                # result = mail.uid('COPY', email_id, processed_mailbox)
                # if result[0] == 'OK':
                #     # Mark the original email for deletion
                #     mail.uid('STORE', email_id, '+FLAGS', '\\Deleted')
                #     mail.expunge()

                # # Add the processed email ID to the set
                # processed_ids.add(email_id.decode())

                # Move the email to the "processed" mailbox
                result = mail.uid('COPY', email_id, processed_mailbox)
                if result[0] == 'OK':
                    # Mark the original email for deletion
                    try:
                        store_result = mail.uid('STORE', email_id, '+FLAGS', '\\Seen')
                        if store_result[0] == 'OK':
                            mail.expunge()
                        else:
                            logging.warning("Failed to mark email as read: %s", email_id)
                    except Exception as e:
                        logging.error("Error marking email as read: %s", email_id, exc_info=True)
                    # Add the processed email ID to the set
                    processed_ids.add(email_id.decode())
                else:
                    logging.warning("Failed to move email to processed mailbox: %s", email_id)

                if file_paths:
                    try:
                        if signature == None:
                            QA.send_qa_to_work_email(file_paths[1], file_paths[0], date_range, email_address)
                            logging.info("Sent with my default signature")
                            send_email_function.send_email(f'QA Sending Success', 'lazaro.gonzalez@conduent.com', f'Email Succesfully Sent at {datetime.now()} to {email_address} sent by {email_sender}')
                        else:
                            QA.send_qa_to_work_email(file_paths[1], file_paths[0], date_range, email_address, signature)
                            logging.info(f'sent with {signature}')
                            send_email_function.send_email(f'QA Sending Success', 'lazaro.gonzalez@conduent.com', f'Email Succesfully Sent at {datetime.now()} to {email_address} sent by {email_sender}')
                    except Exception as e:
                        error_message = str(e)  # Get the error message as a string
                        traceback_str = traceback.format_exc()  # Get the traceback as a string
                        error_message_with_traceback = f"{error_message}\n\nTraceback:\n{traceback_str}"
                        send_email_function.send_email(f'Error QA {email_address}', 'lazaro.gonzalez@conduent.com', error_message_with_traceback)
                        logging.error('Failed to send QA email.', exc_info=True)
                else:
                    send_email_function.send_email(f'QA Missing File', 'lazaro.gonzalez@conduent.com', f'QA Missing File {datetime.now()}')
                    logging.info("No file attachments found.")
                    print("No file attachments found.")

                

# Close the mailbox and log out
mail.close()
mail.logout()

# Log script completion
logging.info("Script execution completed.")
print("Script execution completed.")

