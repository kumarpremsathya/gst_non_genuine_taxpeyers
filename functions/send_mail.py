
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


def send_email(subject, message):
    """
    Send an email notification for manual intervention.
    This function sends an email using Gmail's SMTP server with a predefined sender and receiver.
    The email is formatted to indicate manual intervention is required for a specific subject.
    Parameters:
        subject (str): The subject matter requiring manual intervention. This will be added to the email subject line.
        message (str): The main content of the email. Will be converted to string if not already.
    Returns:
        None
    Raises:
        No exceptions are raised as they are caught and silently passed.
    Note:
        - Uses Gmail SMTP server on port 587
        - Both sender and receiver are set to 'probepoc2023@gmail.com'
        - Requires valid Gmail credentials (password is an app-specific password)
        - Email subject is formatted as "Manual intervention required for {subject}"
    """
    try:
        # Email configuration
        sender_email = 'probepoc2023@gmail.com'
        receiver_email = 'probepoc2023@gmail.com'
        password = 'rovqljwppgraopla'  
      
        # Email content
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = receiver_email
        msg['Subject'] = f"Manual intervention required for {subject}"
        msg.attach(MIMEText(str(message), 'plain'))
        # Connect to the SMTP server
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(sender_email, password)
            server.send_message(msg)
    except Exception as e:
        pass
    