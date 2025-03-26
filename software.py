import os
import re
import time
import random
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

# File Paths
file_path = "usaemails.xlsx"
sent_emails_file = "sent_emails.txt"

# Check if file exists
if not os.path.exists(file_path):
    print(f"Error: File '{file_path}' not found!")
    exit()

# Load Excel File
df = pd.read_excel(file_path)

# Dictionary for common domain corrections
domain_corrections = {
    "samec": "samechemicals.com",
    "pure-chemical": "pure-chemical.com",
    "pqchemicals": "pqchemicals.com",
    "jkenterprises": "jkenterprises.com"
}

# Function to clean and validate emails
def clean_email(email):
    if isinstance(email, str):
        email = re.sub(r"\s+", "", email)  # Remove spaces
        email = re.sub(r"[©#&*^_]", "@", email)  # Replace special symbols with '@'
        
        for key, corrected_domain in domain_corrections.items():
            if key in email:
                email = re.sub(rf"{key}.*", corrected_domain, email)

        email = re.sub(r"(@[a-zA-Z0-9.-]*chemicals)(?!\.com)", r"\1.com", email)

        match = re.search(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", email)
        return match.group(0) if match else None
    return None

# Extract and clean emails, skipping invalid ones
df["Cleaned Email"] = df["Item"].dropna().apply(clean_email)
df = df.dropna(subset=["Cleaned Email"])  # Remove rows with invalid emails

# Load already sent emails
sent_emails = set()
if os.path.exists(sent_emails_file):
    with open(sent_emails_file, "r") as f:
        sent_emails = set(f.read().splitlines())

# Filter unsent emails
df = df[~df["Cleaned Email"].isin(sent_emails)]

# SMTP Configuration
SMTP_USER = "nayachemicalseo@gmail.com"
SMTP_PASSWORD = "vyoscibrerovlzuu"
SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 587

# Image Paths
images = ["p3.jpeg", "p2.jpg"]

# Check if images exist
for img in images:
    if not os.path.exists(img):
        print(f"Error: Image '{img}' not found!")
        exit()

# Email Subject
email_subject = "Naya Chemical New Promotion Items"

# Create SMTP session
try:
    server = smtplib.SMTP(SMTP_HOST, SMTP_PORT)
    server.starttls()
    server.login(SMTP_USER, SMTP_PASSWORD)

    for index, row in df.iterrows():
        try:
            serial_number = row["#"]  # Get serial number
            email = row["Cleaned Email"]
            selected_image = images[index % 2]  # Alternate images

            msg = MIMEMultipart()
            msg["From"] = SMTP_USER
            msg["To"] = email
            msg["Subject"] = email_subject

            # Email Body (HTML)
            html_content = f"""
            <html>
            <body>
                <img src="cid:promo_image" width="600">
                <p>Best Regards,<br>Naya Chemical Team</p>
            </body>
            </html>
            """
            msg.attach(MIMEText(html_content, "html"))

            # Attach Image
            with open(selected_image, "rb") as img_file:
                img = MIMEImage(img_file.read())
                img.add_header("Content-ID", "<promo_image>")
                img.add_header("Content-Disposition", "inline", filename=selected_image)
                msg.attach(img)

            # Send Email
            server.sendmail(SMTP_USER, email, msg.as_string())
            print(f"(# {serial_number}) Email sent to {email} with {selected_image}")

            # Save email to the sent list
            with open(sent_emails_file, "a") as f:
                f.write(email + "\n")

            # **Delay to prevent spam detection**
            wait_time = random.uniform(10, 15)  # Random delay between 10 to 30 seconds
            print(f"Waiting {wait_time:.2f} seconds before sending the next email... ")
            time.sleep(wait_time)

            # **Batch processing** - Reconnect after every 10 emails
            if (index + 1) % 10 == 0:
                server.quit()
                time.sleep(random.uniform(60, 120))  # Wait 1-2 minutes before reconnecting
                server = smtplib.SMTP(SMTP_HOST, SMTP_PORT)
                server.starttls()
                server.login(SMTP_USER, SMTP_PASSWORD)

        except Exception as e:
            print(f"(# {serial_number}) Failed to send email to {email}: {e}")

    server.quit()
    print("All emails sent successfully!")

except Exception as e:
    print(f"SMTP connection error: {e}")
