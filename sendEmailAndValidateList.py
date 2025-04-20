import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os
import pandas as pd
import re
import dns.resolver
import time

# Your email credentials
EMAIL = "ravitejac255@gmail.com"
PASSWORD = ""#app password
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

# Path to your resume
RESUME_PATH = r"" # add resume path 

def sanitize_email(email):
    """Remove any period at the end of the email address."""
    email = str(email)  # Ensure email is a string
    if email.endswith('.'):
        email = email[:-1]
    return email

def is_valid_email_format(email):
    """Check if the email format is valid using regex."""
    email_regex = r'^[a-zA-Z0-9_.+-]+@[a-zAZ0-9-]+\.[a-zA-Z0-9-.]+$'
    return re.match(email_regex, email) is not None

def is_valid_email_domain(email):
    """Check if the email domain can receive emails by looking for MX records."""
    domain = email.split('@')[-1]
    try:
        # Perform DNS lookup for MX records
        dns.resolver.resolve(domain, 'MX')
        return True
    except dns.resolver.NoAnswer:
        print(f"Domain {domain} does not have MX records (cannot receive emails).")
        return False
    except dns.resolver.NXDOMAIN:
        print(f"Domain {domain} is invalid or does not exist.")
        return False

def validate_email(email):
    """Validate the email format and domain before sending email."""
    if pd.isna(email) or email == "":
        print(f"Skipping invalid or empty email: {email}")
        return False

    email = sanitize_email(email)  # Remove trailing period if any

    if not is_valid_email_format(email):
        print(f"The email address {email} is not valid in terms of format.")
        return False

    if not is_valid_email_domain(email):
        print(f"The email address {email} has an invalid or non-existent domain.")
        return False

    print(f"The email address {email} is valid and the domain can receive emails.")
    return True

def send_email(to_email, subject, body):
    """Send email with attachment."""
    msg = MIMEMultipart()
    msg['From'] = EMAIL
    msg['To'] = to_email
    msg['Subject'] = subject

    # Add the message body
    msg.attach(MIMEText(body, 'plain'))

    # Attach the resume
    with open(RESUME_PATH, "rb") as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(RESUME_PATH))
        msg.attach(part)

    # Set up the server
    server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    server.starttls()
    server.login(EMAIL, PASSWORD)

    try:
        # Send the email
        server.sendmail(EMAIL, to_email, msg.as_string())
        print(f"Email sent to {to_email}")
    except smtplib.SMTPRecipientsRefused as e:
        print(f"Error: The email address {to_email} is invalid or rejected: {e}")
    except smtplib.SMTPException as e:
        print(f"Error occurred while sending email to {to_email}: {e}")
    finally:
        server.quit()

def main():
    # Read the Excel file containing email IDs
    file_path = r""  # Path to your Excel file
    df = pd.read_excel(file_path)  # Read the file

    # Check if 'Email' column exists
    if 'Email' not in df.columns:
        print("The Excel file must have a column named 'Email'.")
        return

    # Extract email IDs from the 'Email' column
    email_list = df['Email'].tolist()

    # Email subject and body
    subject = ""
    body = """\
    Hi,
    
    I hope this message finds you well.

    I am reaching out to share my resume and express my interest in potential job opportunities with your organization. 
    
    I have honed my skills in designing, developing, and deploying software solutions that drive business success. I believe my experience, combined with my passion for technology, makes me an ideal candidate to contribute to your team's success.

    Please find my resume attached for your reference. I would greatly appreciate the opportunity to connect and discuss how I can bring value to your organization. I am open to any opportunities you may have and would love to explore the next steps.

    Thank you for considering my application. I look forward to hearing from you soon.

    Thanks & Regards,
    """

    # Initialize counter for email sending
    email_count = 0
    sent_email_ids = []  # List to store email IDs to which the email has been sent

    # Loop through the email list and send emails
    for email_id in email_list:
        if validate_email(email_id):
            send_email(email_id, subject, body)
            sent_email_ids.append(email_id)  # Add email ID to the list of sent emails
            email_count += 1

            # Pause for 10 seconds after every 10 emails
            if email_count % 10 == 0:
                print("Waiting for 10 seconds...")
                time.sleep(10)

    # Print the list of email IDs where emails have been sent
    print("\nEmails have been sent to the following addresses:")
    for sent_email in sent_email_ids:
        print(sent_email)

if __name__ == "__main__":
    main()
