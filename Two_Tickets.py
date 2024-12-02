import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

# Load Excel sheet
file_path = r"C:\Users\tatho\Desktop\Tickets Info\Duo Tickets 2.0.xlsx"
data = pd.read_excel(file_path)

# Font setup (adjust path to match your system)
FONT_PATH = r"C:\Windows\Fonts\GARABD.TTF"
FONT_SIZE = 30

# Email setup
SMTP_SERVER = "smtp.gmail.com"  # Adjust to your SMTP server
SMTP_PORT = 587
EMAIL_USER = "swcpromcommittee@gmail.com"  # Replace with your email
EMAIL_PASS = "ijbr jqsa vtak xhkq"  # Replace with your email password
DEFAULT_FROM_EMAIL = 'swcpromcommittee@gmail.com'

# Ensure output directory exists
os.makedirs("generated_images", exist_ok=True)

# Generate images and send emails
for index, row in data.iterrows():
    # Extract data from the row
    first_name = row.iloc[2]  # Column C (index 2)
    last_name = row.iloc[3]  # Column D (index 3)
    guest_name_1 = row.iloc[6]  # Column G (index 6)
    your_email = row.iloc[4]  # Column E (index 4)
    cbe_email_1 = row.iloc[7]  # Column H (index 7)

    guest_name_1 = guest_name_1 if pd.notnull(guest_name_1) else ""

    # Load your pre-designed image
    template_image_path = r"C:\Users\tatho\Downloads\formal ticket template.png"  # Replace with your template's path
    template_img = Image.open(template_image_path)

    # Ensure the template has an RGBA mode (if you want transparency)
    template_img = template_img.convert("RGBA")

    # Create a drawable canvas on top of the image
    draw = ImageDraw.Draw(template_img)

    # Define text properties
    FONT_SIZE = 70  # Adjust this to make the font bigger
    LINE_SPACING = 30  # Adjust this for more space between lines
    text = f"{first_name} {last_name}\n{guest_name_1}"
    font = ImageFont.truetype(FONT_PATH, FONT_SIZE)

    # Calculate the total height of the text block, including line spacing
    lines = text.split('\n')
    text_width = max(draw.textbbox((0, 0), line, font=font)[2] for line in lines)
    text_height = sum(draw.textbbox((0, 0), line, font=font)[3] for line in lines) + (LINE_SPACING * (len(lines) - 1))

    # Calculate the position to center the text block
    position_x = (template_img.width - text_width) // 2
    position_y = (template_img.height - text_height) // 2

    # Overlay each line of text with spacing
    current_y = position_y
    for line in lines:
        line_width, line_height = draw.textbbox((0, 0), line, font=font)[2:4]
        line_x = (template_img.width - line_width) // 2  # Center each line
        draw.text((line_x, current_y), line, fill="white", font=font, align="center")
        current_y += line_height + LINE_SPACING  # Move to the next line position

    # Save the modified image
    image_path = f"generated_images/ticket_{your_email}.png"
    template_img.save(image_path)

    # Send email with the image
    # Generate the list of recipients, ensuring each email is a string and valid
    recipients = [
        str(your_email) if pd.notnull(your_email) and your_email != "" else None,
        str(cbe_email_1) if pd.notnull(cbe_email_1) and cbe_email_1 != "" else None,
    ]

    # Remove any None values (invalid email addresses)
    recipients = [email for email in recipients if email]

    # Debugging: Print recipients to ensure the list is correct
    print(f"Recipients for index {index}: {recipients}")

    # If no valid recipients are found, skip sending this email
    if not recipients:
        print(f"Skipping row {index} due to missing or invalid email addresses.")
        continue

    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText

    # Email setup
    msg = MIMEMultipart()
    msg["From"] = EMAIL_USER
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = "🎟 Your Ticket + Event Details"

    # HTML Body
    body = """
    <html>
    <body>
        <p>Hi {first_name},</p>
        <p>Your ticket is <b>attached</b> to this email. Please <b>save it on your phone</b> and <b>share it with your guests.</b> All attendees must show their <b>ticket</b> and <b>School ID</b> at the entrance.</p>

        <h3><u>Event Details<u></h3>
        <ul>
            <li><b>Date:</b> Wednesday, December 4th</li>
            <li><b>Time:</b> 6:00 PM - 8:00 PM (<i>Doors close at 6:30 PM</i>)</li>
            <li><b>Location:</b> SWC Gym</li>
            <li><b>Included:</b> Snacks + Free Photo Booth Prints</li>
        </ul>

        <h3><u>Event Guidelines<u></h3>
        <ul>
            <li><b>Bring:</b> School ID + Phone</li>
            <li><b>Guests:</b> Must have the ticket and be an SWC student</li>
            <li><b>Entrance Requirements:</b></li>
            <ul>
                <li>SWC students only</li>
                <li>Bag & coat checks at the entrance by Administration</li>
                <li><b>No re-entry after 6:30 PM</b></li>
            </ul>
            <li><b>Prohibited:</b></li>
            <ul>
                <li><b>NO outside food or drink</b></li>
                <li><b>NO alcohol or illicit substances</b></li>
            </ul>
        </ul>

        <p><b>Reminder:</b> All attendees must be SWC students.</p>

        <p><i>In case of emergency, contact +1(403)-383-9086 or a teacher supervisor.<i></p>

        <p>Thank you, and we can’t wait to see you there!</p>
        <p><b>Winter Formal Organizing Committee</b></p>
    </body>
    </html>
    """.format(first_name=first_name)

    # Attach the HTML body
    msg.attach(MIMEText(body, "html"))

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
