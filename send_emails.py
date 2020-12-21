"""
Sends email from a csv list with attached image
using a GMAIL account

It assumes that the name of the image to attach
has the name of the person in the CSV
"""

import smtplib
import ssl
import csv

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage


sender_email = "myemail@gmail.com"
sender_password = "PASSWORD"

from_email = "myemail@gmail.com"
subject = "Subject"
filename = "Image Name.jpg"
file_path = "<list and images directory path>"


def send_emails():
    with open(file_path + 'Listado.csv') as csv_file:
        csv_reader = csv.DictReader(csv_file)
        line_count = 0
        for row in csv_reader:
            line_count += 1
            try:
                name = row.get("name", "").strip().title()
                if name:
                    email = row.get("email", "")
                    if email:
                        send_email(sender_password, email.strip(), name + ".jpg")
                    else:
                        print("No Email - LINE " + str(line_count) + " Name: " + name)
            except Exception as e:
                print(f"Unknown Error: LINE {line_count}", e)


def send_email(sender_password, receiver_email, local_filename):
    message = MIMEMultipart("alternative")
    message["Subject"] = subject
    message["From"] = from_email
    message["To"] = receiver_email

    # Create the plain-text and HTML version of your message
    text = """\
    Text email body"""

    html = """\
    <html>
      <body>
        <p>HTML email body</p>
        <img src="cid:attached_image" style="max-width: 500px; height: auto">
      </body>
    </html>
    """

    # Turn these into plain/html MIMEText objects
    part1 = MIMEText(text, "plain")
    part2 = MIMEText(html, "html")

    # Add HTML/plain-text parts to MIMEMultipart message
    # The email client will try to render the last part first
    message.attach(part1)
    message.attach(part2)

    local_file = file_path + local_filename

    # Open PDF file in binary mode
    print(f"Uplading file {local_file}")
    with open(local_file, "rb") as attachment:
        part = MIMEImage(attachment.read(), _subtype="jpeg")

    # Add header as key/value pair to attachment part
    part.add_header('Content-ID', '<attached_image>')
    part.add_header(
        "Content-Disposition",
        f"attachment; filename={filename}",
    )

    # Add attachment to message and convert message to string
    message.attach(part)

    # Create secure connection with server and send email
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(sender_email, sender_password)
        print(f"Sending email to {receiver_email}")
        server.sendmail(
            from_email, receiver_email, message.as_string()
        )
        print("Email Sent")
