import boto3
import os
import datetime

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

def send_email_with_attachment(Email, file_name, file_path):
    # Replace these values with your own
    x = datetime.datetime.now()
    sender_email =  <add sender's email ID which is verified in ses>
    recipient_email = Email
    email_subject = "AWS Security Hub Report"
    email_body = 'HELLO, \n\n PFA  AWS Security Hub Report generated on '+ x.strftime("%c") + ' UTC'
    attachment_path = file_path  # Lambda has /tmp directory for write access
    aws_region = 'us-east-1'

    # Create MIME object for the email
    msg = MIMEMultipart()
    msg['Subject'] = email_subject
    msg['From'] = sender_email
    msg['To'] = recipient_email

    # Attach the body of the email
    msg.attach(MIMEText(email_body, 'plain'))

    # Attach the file
    with open(attachment_path, 'rb') as attachment:
        part = MIMEApplication(attachment.read(), Name=file_name)
        part['Content-Disposition'] = f'attachment; filename={os.path.basename(attachment_path)}'
        msg.attach(part)

    # Convert the message to a string
    raw_message = msg.as_string()

    # Send the email
    ses = boto3.client('ses', region_name=aws_region)
    try:
        response = ses.send_raw_email(
            Source=sender_email,
            Destinations=[recipient_email],
            RawMessage={'Data': raw_message}
        )
        print(f"Email sent! Message ID: {response['MessageId']}")

    except Exception as e:
        print(f"Error: {e}")

