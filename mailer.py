import os
from dotenv import load_dotenv
import win32com.client as client
import csv


class Mailer:
    def __init__(
        self,
        recipient_email,
        cc_email,
        subject,
        recipient_name,
        template_html,
        attachment_path,
    ):
        self.recipient_email = recipient_email
        self.cc_email = cc_email
        self.subject = subject
        self.recipient_name = recipient_name
        self.template_html = template_html
        self.attachment_path = attachment_path

    def create_mail(self, display=False, send=False):
        outlook = client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)

        if display:
            mail.Display()

        mail.Attachments.Add(self.attachment_path)
        mail.To = self.recipient_email
        mail.CC = self.cc_email
        mail.Subject = self.subject
        mail.HTMLBody = self.template_html.format(self.recipient_name)
        
        if send:
            mail.Send()
        else:
            mail.Save()

        print(f"Email dispatched to {self.recipient_email} with attachment: {self.attachment_path}")


class RecipientList:
    def __init__(
        self, recipients_path, output_path, output_file_name_template, debug=False
    ):
        self.recipients = []
        self.recipients_path = recipients_path
        self.output_path = output_path
        self.output_file_name_template = output_file_name_template
        self.debug = debug

    def get_recipients(self):
        with open(self.recipients_path) as csv_file:
            csv_reader = csv.reader(csv_file, delimiter=",")
            page_index = 0

            for row in csv_reader:
                self.add_recipient(page_index, row)
                page_index += 1

            print(f"Processed {page_index} recipient(s) for mailing")

        return self.recipients

    def add_recipient(self, page_index, row):
        recipient_name = row[0]
        recipient_name_stub = "_".join(recipient_name.split())
        first_name = recipient_name.split()[0]

        file_name = self.output_file_name_template.format(
            index=page_index + 1, recipient=recipient_name_stub
        )

        email = row[1].strip()
        attachment_rel_path = f"{self.output_path}\\{file_name}.pdf"
        attachment_abs_path = os.path.abspath(attachment_rel_path)

        self.recipients.append(
            {
                "first_name": first_name,
                "full_name": recipient_name,
                "email": email,
                "attachment": attachment_abs_path,
            }
        )

        if self.debug:
            print(
                f"Processed Recipient=({first_name}, {recipient_name}, {email}, {attachment_abs_path})"
            )


class TemplateReader:
    def __init__(self, template_path):
        self.template_path = template_path

    def get_template(self):
        print(f"Using template for email: {self.template_path}")

        with open(self.template_path) as file:
            template_html = file.read()

        return template_html


if __name__ == "__main__":
    load_dotenv()

    recipients_list = RecipientList(
        recipients_path=os.getenv("RECIPIENTS_PATH"),
        output_path=os.getenv("OUTPUT_PATH"),
        output_file_name_template=os.getenv("OUTPUT_FILE_NAME_TEMPLATE"),
    )

    recipients = recipients_list.get_recipients()

    template_html = TemplateReader(os.getenv("EMAIL_TEMPLATE_PATH")).get_template()
    cc_email = os.getenv("CC_EMAIL")
    subject = os.getenv("SUBJECT")

    for recipient in recipients:

        mailer = Mailer(
            recipient_email=recipient["email"],
            recipient_name=recipient["first_name"],
            attachment_path=recipient["attachment"],
            cc_email=cc_email,
            subject=subject,
            template_html=template_html,
        )

        mailer.create_mail(send=False)