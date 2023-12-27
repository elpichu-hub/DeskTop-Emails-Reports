import email
import imaplib
import re
import logging
import traceback
import quopri
from datetime import datetime
import email_config
import send_email_function
from gpt_email_responder_images import call_gpt_images
import chardet

# Set up logging
log_filename = "chatgptImagesAssistant.log"
logging.basicConfig(filename=log_filename, level=logging.INFO, format='%(asctime)s %(levelname)s: %(message)s')

# Set up your email and password
username = email_config.EMAIL_ADDRESS_AUTO
password = email_config.EMAIL_PASSWORD_AUTO

# Set the mailbox you want to check
inbox_mailbox = "inbox"
processed_mailbox = "processed"

# Set the keyword you want to search for in the subject
keyword = 'IMAGES'

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


            content_type = email_message.get('Content-Type', '')
            match = re.search(r'charset="([^"]+)"', content_type)
            if match:
                encoding = match.group(1)
            else:
                encoding = chardet.detect(email_body)['encoding']
            email_body = quopri.decodestring(email_body).decode(encoding, errors='replace')
            
            # Remove mailto links
            email_body = re.sub(r'<mailto:[^>]*>', '', email_body)

            # Remove the signature
            signature_index = email_body.find("Lazaro Gonzalez")
            if signature_index != -1:
                email_body = email_body[:signature_index]
            
            # Move the email to the "processed" mailbox
            result = mail.uid('COPY', email_id, processed_mailbox)
            if result[0] == 'OK':
                # Mark the original email for deletion
                mail.uid('STORE', email_id, '+FLAGS', '\\Deleted')
                mail.expunge()

            # Get the email sender
            email_sender = email.utils.parseaddr(email_message['From'])[1]
            logging.info("Email sender: %s", email_sender)
            print("Email sender: %s", email_sender)

            if email_body:
                try:
                    gpt_response = call_gpt_images(email_body, email_sender)
                    logging.info("Response: %s", email_body)
                    logging.info(f'ChatGPT-IMAGES, succesful')
                except Exception as e:
                    error_message = str(e)  # Get the error message as a string
                    traceback_str = traceback.format_exc()  # Get the traceback as a string
                    error_message_with_traceback = f"{error_message}\n\nTraceback:\n{traceback_str}"
                    send_email_function.send_email(f'ChatGPT-IMAGES', 'lazaro.gonzalez@conduent.com', str(error_message_with_traceback))
                    send_email_function.send_email(f'ChatGPT-IMAGES', email_sender, str(error_message_with_traceback))
                    logging.error('ChatGPT-IMAGES.', exc_info=True)
            else:
                send_email_function.send_email(f'No ChatGPT-IMAGES-REQUEST', 'lazaro.gonzalez@conduent.com', f'ChatGPT-IMAGES {datetime.now()}')
            

# Close the mailbox and log out
mail.close()
mail.logout()

# Log script completion
logging.info("Script execution completed.")
print("Script execution completed.")
