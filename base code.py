import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

# Load Excel sheet
file_path = r"C:\Users\tatho\Documents\Tickets Demo System.xlsx"
data = pd.read_excel(file_path)

# Font setup (adjust path to match your system)
FONT_PATH = r"C:\Windows\Fonts\Arial.ttf"
FONT_SIZE = 30

# Email setup
SMTP_SERVER = "smtp.gmail.com"  # Adjust to your SMTP server
SMTP_PORT = 587
EMAIL_USER = "swcpromcommittee@gmail.com"  # Replace with your email
EMAIL_PASS = "ijbr jqsa vtak xhkq"  # Replace with your email password
DEFAULT_FROM_EMAIL='swcpromcommittee@gmail.com'

# Ensure output directory exists
os.makedirs("generated_images", exist_ok=True)

# Generate images and send emails
for index, row in data.iterrows():
    # Extract data from the row
    first_name = row.iloc[2]  # Column C (index 2)
    last_name = row.iloc[3]  # Column D (index 3)
    guest_name_1 = row.iloc[6]  # Column G (index 6)
    guest_name_2 = row.iloc[8]  # Column I (index 8)
    your_email = row.iloc[4]  # Column E (index 4)
    cbe_email_1 = row.iloc[7]  # Column H (index 7)
    cbe_email_2 = row.iloc[9]  # Column J (index 9)

    # Create the text for the image
    text = f"SWC Winter Formal 2024\nThis ticket is valid for\n{first_name} {last_name}\n{guest_name_1}\n{guest_name_2}"

    # Generate the image
    img = Image.new("RGB", (800, 500), color="black")  # Adjust dimensions as needed
    draw = ImageDraw.Draw(img)
    font = ImageFont.truetype(FONT_PATH, FONT_SIZE)
    text_width, text_height = draw.textbbox((0, 0), text, font=font)[2:]
    position = ((800 - text_width) // 2, (500 - text_height) // 2)
    draw.multiline_text(position, text, fill="white", font=font, align="center")

    # Save the image
    image_path = f"generated_images/ticket_{index}.png"
    img.save(image_path)

    # Send email with the image
    # Generate the list of recipients, ensuring each email is a string and valid
    recipients = [
        str(your_email) if pd.notnull(your_email) and your_email != "" else None,
        str(cbe_email_1) if pd.notnull(cbe_email_1) and cbe_email_1 != "" else None,
        str(cbe_email_2) if pd.notnull(cbe_email_2) and cbe_email_2 != "" else None
    ]

    # Remove any None values (invalid email addresses)
    recipients = [email for email in recipients if email]

    # Debugging: Print recipients to ensure the list is correct
    print(f"Recipients for index {index}: {recipients}")

    # If no valid recipients are found, skip sending this email
    if not recipients:
        print(f"Skipping row {index} due to missing or invalid email addresses.")
        continue

    msg = MIMEMultipart()
    msg["From"] = EMAIL_USER
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = "Your Ticket (Demo)"
    body = "Please find your ticket attached. not the real thing bro"
    msg.attach(MIMEText(body, "plain"))

    # Attach the image
    with open(image_path, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition",
        f"attachment; filename={os.path.basename(image_path)}",
    )
    msg.attach(part)

    # Send the email
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(EMAIL_USER, EMAIL_PASS)
        server.send_message(msg)

print("All images generated and emails sent!")
