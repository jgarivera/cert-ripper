import json
import os


class RecipientList:
    def __init__(
        self,
        start_page_index: int = 0,
        json_points_path: str = None,
        ripped_certs_path: str = None,
        ripped_cert_file_name: str = None,
        debug: bool = False
    ) -> None:
        self.start_page_index = start_page_index
        self.json_points_path = json_points_path
        self.ripped_certs_path = ripped_certs_path
        self.ripped_cert_file_name = ripped_cert_file_name
        self.recipients = []
        self.debug = debug

    def get_recipients(self) -> list:
        current_page_index = self.start_page_index
        total_recipients = 0

        with open(self.json_points_path) as json_file:
            points = json.load(json_file)

            for point in points:
                point_name: str = point["name"]
                point_tag: str = point["tag"]
                point_recipients: list = point["recipients"]

                for point_recipient in point_recipients:
                    self.add_recipient(
                        current_page_index,
                        point_name,
                        point_tag,
                        point_recipient
                    )

                    current_page_index += 1
                    total_recipients += 1

            print(
                f"Read \x1b[95m{total_recipients} recipient(s)\x1b[0m for mailing")

        return self.recipients

    def add_recipient(self, page_index: int, point_name: str, point_tag: str, point_recipient: dict[str, str]) -> None:
        certificate = point_name
        recipient_name = point_recipient["name"]
        recipient_email = point_recipient["email"]
        recipient_first_name = self.get_first_name(recipient_name)
        attachment_path = self.get_attachment(
            page_index, point_tag, recipient_name)

        self.recipients.append(
            {
                "certificate": certificate,
                "first_name": recipient_first_name,
                "full_name": recipient_name,
                "email": recipient_email,
                "attachment": attachment_path,
            }
        )

        if self.debug:
            print(
                f"Added MailRecipient=({recipient_first_name}, {certificate}, {recipient_name}, {recipient_email}, {attachment_path})"
            )

    def get_first_name(self, full_name: str) -> str:
        return full_name.split()[0].strip()

    def get_attachment(self, page_index: int, tag: str, recipient_name: str) -> str:
        recipient_name_slug = "_".join(recipient_name.split())

        file_name = self.ripped_cert_file_name.format(
            index=page_index + 1, tag=tag, recipient=recipient_name_slug
        )

        return os.path.abspath(f"{self.ripped_certs_path}\\{file_name}.pdf")
