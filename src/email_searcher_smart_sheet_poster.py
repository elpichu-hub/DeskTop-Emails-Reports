import email
import imaplib
import os
import email_config
from datetime import datetime
import logging
import send_email_function
import traceback
import encrypting
import smart_sheet_poster


# Set up logging
log_filename = "smart_sheet_poster.log"
logging.basicConfig(filename=log_filename, level=logging.INFO, format='%(asctime)s %(levelname)s: %(message)s')


# save the attachments' path from the email to the pass it to a function
def save_attachments(email_message, sub_folder_path):
    file_paths = []

    # If the email has attachments
    for part in email_message.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue

        filename = part.get_filename()

        if bool(filename):
            # Only save .xlsx, .xls and .csv attachments
            if not (filename.lower().endswith('.xlsx') or filename.lower().endswith('.xls') or filename.lower().endswith('.csv')):
                continue
            file_path = os.path.join(sub_folder_path, filename)

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

    return file_paths


# Set up your email and password
username = email_config.EMAIL_ADDRESS_AUTO
password = email_config.EMAIL_PASSWORD_AUTO

# Set the mailbox you want to check
inbox_mailbox = "inbox"
processed_mailbox = "processed"

# Set the keyword you want to search for in the subject
keyword = 'Poster'

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
email_address = None
Signature = None


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

                # Create a subfolder with the current time
                sub_folder_name = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
                sub_folder_path = os.path.join(folder_name, sub_folder_name)
                if not os.path.exists(sub_folder_path):
                    os.makedirs(sub_folder_path)

                file_paths = save_attachments(email_message, sub_folder_path)

                # Move the email to the "processed" mailbox
                result = mail.uid('COPY', email_id, processed_mailbox)
                if result[0] == 'OK':
                    # Mark the original email for deletion
                    mail.uid('STORE', email_id, '+FLAGS', '\\Deleted')
                    mail.expunge()

                # Add the processed email ID to the set
                processed_ids.add(email_id.decode())


                target_email = 'lazarogonzalez.auto@gmail.com'

                if file_paths:
                    
                    try:
                        # Decrypt the file received by email with subject line 'Poster'
                        encrypting.main_func_decrypt(file_paths[0], email_config.encryption_key)
                        # Send an email to the target email address with confirmation that the file was decrypted
                        send_email_function.send_email(f'Worked', target_email, f'Worked Succesfully at {datetime.now()}')
                        # logg the file path and email address
                        logging.info(f'File {file_paths[0]}, email {email_address}')

                        # Post the file to SmartSheet
                        try:
                            smart_sheet_poster.smart_sheet_poster_function('decrypted_lookups.csv')
                            logging.info(f'Posted Succesfully at {datetime.now()}')
                        except Exception as e:
                            error_message = str(e)
                            traceback_str = traceback.format_exc()
                            error_message_with_traceback = f"{error_message}\n\nTraceback:\n{traceback_str}"
                            send_email_function.send_email(f'Error Poster', target_email, error_message_with_traceback)
                            logging.error('Failed to send Poster email.', exc_info=True)

                    except Exception as e:
                        error_message = str(e)  # Get the error message as a string
                        traceback_str = traceback.format_exc()  # Get the traceback as a string
                        error_message_with_traceback = f"{error_message}\n\nTraceback:\n{traceback_str}"
                        send_email_function.send_email(f'Error Poster', target_email, error_message_with_traceback)
                        logging.error('Failed to send Poster email.', exc_info=True)
                else:
                    # Send an email to the target email address with confirmation that the file was not decrypted due to missing file
                    send_email_function.send_email(f'Poster Missing File', target_email, f'Poster Missing File {datetime.now()}')
                    logging.info("No file attachments found.")
                    print("No file attachments found.")

                
# Close the mailbox and log out
mail.close()
mail.logout()

# Log script completion
logging.info("Script execution completed.")
print("Script execution completed.")

