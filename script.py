import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from dotenv import load_dotenv
import os

load_dotenv()

# Function to send email with attachment


def send_email(to_email, subject, body, from_email, password, resume_path):
    try:

        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)

        server.login(from_email, password)

        # Construct the email
        msg = MIMEMultipart()
        msg['From'] = from_email
        msg['To'] = to_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        # Attach the resume
        with open(resume_path, "rb") as f:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition',
                            f'attachment; filename="resume.pdf"')
            msg.attach(part)

        # Send the email
        server.send_message(msg)
        print(f"Email sent to {to_email}")
        server.quit()

    except Exception as e:
        print(f"Failed to send email to {to_email}. Error: {str(e)}")

# Load email addresses from Excel sheet


def load_emails_from_excel(file_path):
    try:
        data = pd.read_excel(file_path)
        if 'Email' in data.columns:
            return data['Email'].tolist()
        else:
            raise KeyError(
                "The 'Email' column was not found in the Excel file.")
    except Exception as e:
        print(f"Error reading the Excel file: {str(e)}")
        return []


# Main function
if __name__ == "__main__":
    # Load email credentials from environment variables
    from_email = os.getenv('EMAIL_USER')
    password = os.getenv('EMAIL_PASS')

    # Load email list from Excel (use raw string to avoid backslash issues)
    file_path = os.getenv('EMAIL_LIST_PATH', 'emails.xlsx')
    email_list = load_emails_from_excel(file_path)

    # Print the loaded email list
    print("Loaded email list:", email_list)

    # Email content
    subject = " Application for Software Developer Role/Internship"
    body_template = """
Dear Hiring Team,

I hope this message finds you well.

I am reaching out to express my interest in joining your team as a software developer or intern. Having followed the impactful work your organization has been doing, I am enthusiastic about the possibility of contributing meaningfully through my skills in both frontend and backend development.

Over the past year, I’ve worked on multiple hands-on projects—ranging from real-time dashboards to user-focused web applications—that have strengthened my problem-solving and collaboration abilities. I'm confident that this foundation, coupled with my eagerness to learn, aligns with the kind of energy and commitment your team values.

Attached is my resume for your review. I’d be honored to explore how I can bring value to your organization and grow alongside your talented team.

Looking forward to the opportunity.

Warm regards,  
Asghar Ali  
Contact: 0319 2583564  
Email: asgharali.224415@gmail.com

    """

    # Path to your resume file (adjust the path as needed)
    # Replace with your actual path
    resume_path = os.getenv('RESUME_PATH')

    # Loop through each email and send
    if email_list:
        for email in email_list:
            send_email(email, subject, body_template,
                       from_email, password, resume_path)
    else:
        print("No emails to send.")
