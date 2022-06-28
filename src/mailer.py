import os
from dotenv import load_dotenv
import win32com.client as client
from utils.recipient_list import RecipientList
from utils.template_reader import TemplateReader


class Mailer:
    def __init__(
        self,
        cc_email: str,
        subject: str,
        recipient_email: str,
        recipient_first_name: str,
        certificate_name: str,
        template_html: str,
        attachment_path: str,
    ):

        self.cc_email = cc_email
        self.subject = subject
        self.recipient_first_name = recipient_first_name
        self.recipient_email = recipient_email
        self.certificate_name = certificate_name
        self.template_html = template_html
        self.attachment_path = attachment_path

    def create_mail(self, display=False, send=False) -> None:
        outlook = client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        certificate = self.certificate_name

        if display:
            mail.Display()

        mail.To = self.recipient_email
        mail.CC = self.cc_email
        mail.Subject = self.subject.format(
            certificate=certificate
        )
        mail.HTMLBody = self.template_html.format(
            recipient=self.recipient_first_name,
            certificate=certificate.lower()
        )
        mail.Attachments.Add(self.attachment_path)

        mail.Save()

        print(
            f"{certificate} email dispatched to \x1b[94m{self.recipient_email}\x1b[0m")


if __name__ == "__main__":
    load_dotenv()

    recipients_list = RecipientList(
        start_page_index=int(os.getenv("START_PAGE_INDEX")),
        json_points_path=os.getenv("JSON_POINTS_PATH"),
        ripped_certs_path=os.getenv("RIPPED_CERTS_PATH"),
        ripped_cert_file_name=os.getenv("RIPPED_CERT_FILE_NAME"),
        debug=True
    )

    recipients = recipients_list.get_recipients()

    template_html = TemplateReader(
        os.getenv("EMAIL_TEMPLATE_PATH")).read()

    cc_email = os.getenv("CC_EMAIL")
    subject = os.getenv("SUBJECT")

    for recipient in recipients:
        mailer = Mailer(
            cc_email=cc_email,
            subject=subject,
            recipient_email=recipient["email"],
            recipient_first_name=recipient["first_name"],
            certificate_name=recipient["certificate"],
            template_html=template_html,
            attachment_path=recipient["attachment"],
        )

        mailer.create_mail(send=False)
