from dotenv import load_dotenv
from PyPDF2 import PdfFileReader, PdfFileWriter
import os
import json


class CertRipper:
    def __init__(
        self,
        start_page_index=0,
        master_pdf_path=None,
        json_points_path=None,
        ripped_certs_path=None,
        ripped_cert_file_name=None,
    ):
        self.start_page_index = start_page_index
        self.master_pdf_path = master_pdf_path
        self.pdf = PdfFileReader(master_pdf_path)
        self.pdf_length = self.pdf.getNumPages()
        self.json_points_path = json_points_path
        self.ripped_certs_path = ripped_certs_path
        self.ripped_cert_file_name = ripped_cert_file_name

    def process(self):
        recipient_groups = self.get_recipient_groups_from_points()
        self.extract_pdf_from_master(recipient_groups)

    def extract_pdf_from_master(self, recipient_groups):
        current_page_index = self.start_page_index
        process_index = 0

        for recipient_group in recipient_groups:
            recipient_group_name = recipient_group["name"]
            recipient_group_tag = recipient_group["tag"]
            recipient_slugs = recipient_group["recipient_slugs"]

            print(
                f"[*] Ripping \x1b[93m{recipient_group_name}\x1b[0m group ...")

            for recipient_slug in recipient_slugs:
                page = self.pdf.getPage(current_page_index)

                file_name = self.ripped_cert_file_name.format(
                    index=current_page_index + 1,
                    tag=recipient_group_tag,
                    recipient=recipient_slug
                )

                pdf_writer = PdfFileWriter()
                pdf_writer.addPage(page)

                output_file_name = f"{self.ripped_certs_path}\\{file_name}.pdf"

                with open(output_file_name, "wb") as out:
                    pdf_writer.write(out)

                print(
                    f"\x1b[95m[{process_index}]\x1b[0m Ripped \x1b[92m[{file_name}]\x1b[0m from \x1b[94mpage {current_page_index + 1}\x1b[0m of master")
                current_page_index += 1
                process_index += 1

    def get_recipient_groups_from_points(self):
        recipient_groups = []
        total_recipients = 0

        with open(self.json_points_path, "r") as json_file:
            points = json.load(json_file)

            for point in points:
                point_name = point["name"]
                point_tag = point["tag"]
                point_recipients = point["recipients"]
                point_recipient_slugs = []

                for point_recipient in point_recipients:
                    recipient_name = point_recipient["name"]
                    recipient_name_slug = "_".join(recipient_name.split())
                    point_recipient_slugs.append(recipient_name_slug)
                    total_recipients += 1

                recipient_groups.append({
                    "name": point_name,
                    "tag": point_tag,
                    "recipient_slugs": point_recipient_slugs
                })

            total_groups = len(recipient_groups)

            self.__check_pdf_length(total_recipients)

            print(
                f"Read \x1b[95m{total_groups} groups(s)\x1b[0m and \x1b[95m{total_recipients} recipient(s)\x1b[0m from JSON points")

        return recipient_groups

    def __check_pdf_length(self, recipients_length):
        pdf_length = self.pdf_length - (self.start_page_index)
        if pdf_length != recipients_length:
            raise ValueError(
                f"Number of recipients ({recipients_length}) does not match with PDF length ({pdf_length})"
            )


if __name__ == "__main__":
    load_dotenv()

    ripper = CertRipper(
        start_page_index=os.getenv("START_PAGE_INDEX"),
        master_pdf_path=os.getenv("MASTER_PDF_PATH"),
        json_points_path=os.getenv("JSON_POINTS_PATH"),
        ripped_certs_path=os.getenv("RIPPED_CERTS_PATH"),
        ripped_cert_file_name=os.getenv("RIPPED_CERT_FILE_NAME"),
    )

    ripper.process()
