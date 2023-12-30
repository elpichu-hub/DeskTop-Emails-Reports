import sys
import os

# Get the current working directory
current_dir = os.getcwd()

# Get the parent directory by moving up one level
parent_dir = os.path.dirname(current_dir)

# Add the parent directory to the sys.path list
sys.path.append(parent_dir)

# Now, you can import a module from the parent directory
import email_config

# sets up email sending function to accept attachments
def send_email(subject, recipient, body, wrappup_file=None):
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

    if wrappup_file is not None:
        with open(wrappup_file, 'rb') as file:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(file.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(wrappup_file))
            message.attach(part)

    # Connect to the Gmail SMTP server and send the email
    with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
        smtp.starttls()
        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        smtp.send_message(message)