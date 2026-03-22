import os
import smtplib
import mimetypes
from email.message import EmailMessage


class EmailSender:
    def __init__(self, smtp_server, smtp_port, sender_email, app_password, sender_name=""):
        self.smtp_server = smtp_server
        self.smtp_port = int(smtp_port)
        self.sender_email = sender_email
        self.app_password = app_password
        self.sender_name = sender_name.strip()

    def create_message(self, recipient_email, subject, body, attachment_path=None):
        msg = EmailMessage()

        if self.sender_name:
            msg["From"] = f"{self.sender_name} <{self.sender_email}>"
        else:
            msg["From"] = self.sender_email

        msg["To"] = recipient_email
        msg["Subject"] = subject
        msg.set_content(body)

        if attachment_path:
            self.add_attachment(msg, attachment_path)

        return msg

    def add_attachment(self, msg, attachment_path):
        if not os.path.exists(attachment_path):
            raise FileNotFoundError(f"Attachment not found: {attachment_path}")

        mime_type, _ = mimetypes.guess_type(attachment_path)
        if mime_type:
            maintype, subtype = mime_type.split("/", 1)
        else:
            maintype, subtype = "application", "octet-stream"

        with open(attachment_path, "rb") as file:
            file_data = file.read()

        msg.add_attachment(
            file_data,
            maintype=maintype,
            subtype=subtype,
            filename=os.path.basename(attachment_path)
        )

    def send_message(self, msg):
        with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
            server.starttls()
            server.login(self.sender_email, self.app_password)
            server.send_message(msg)