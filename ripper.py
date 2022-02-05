from dotenv import load_dotenv
from PyPDF2 import PdfFileReader, PdfFileWriter
import pdfplumber
import csv
import os


class CertRipper:
    CSV_LIST_MODE = "CSV_LIST_MODE"
    PDF_TEXT_MODE = "PDF_TEXT_MODE"

    def __init__(
        self,
        mode=CSV_LIST_MODE,
        start_page_index=0,
        master_pdf_path=None,
        recipients_path=None,
        output_path=None,
        output_file_name_template=None,
    ):
        self.mode = mode
        self.start_page_index = start_page_index
        self.master_pdf_path = master_pdf_path
        self.pdf = PdfFileReader(master_pdf_path)
        self.pdf_length = self.pdf.getNumPages()
        self.recipients_path = recipients_path
        self.output_path = output_path
        self.output_file_name_template = output_file_name_template

    def process(self):
        if self.is_csv_list_mode():
            self.participants = self.get_recipients_from_csv()
            self.extract_pdf_from_master()
            return

        self.pdf_for_text = pdfplumber.open(self.master_pdf_path)
        self.extract_pdf_from_master()

    def extract_pdf_from_master(self):
        for page_index in range(self.start_page_index, self.pdf_length):
            page = self.pdf.getPage(page_index)

            recipient_name_stub = self.extract_recipient_name_stub(page_index)
            file_name = self.output_file_name_template.format(
                index=page_index + 1, recipient=recipient_name_stub
            )

            pdf_writer = PdfFileWriter()
            pdf_writer.addPage(page)

            output_file_name = f"{self.output_path}\\{file_name}.pdf"

            with open(output_file_name, "wb") as out:
                pdf_writer.write(out)

            print("Extracted: {}".format(output_file_name))

    def is_csv_list_mode(self):
        if self.mode is None:
            raise ValueError("Cert ripper mode not specified")

        return self.mode is CertRipper.CSV_LIST_MODE

    def extract_recipient_name_stub(self, page_index):
        if not self.is_csv_list_mode():
            raise ValueError(
                f"Attempted to extract participant name from CSV while in {CertRipper.PDF_TEXT_MODE}"
            )

        return self.participants[self.start_page_index - page_index]

    def get_recipients_from_csv(self):
        recipients = []

        with open(self.recipients_path) as csv_file:
            csv_reader = csv.reader(csv_file, delimiter=",")

            for row in csv_reader:
                recipient_name = row[0]
                recipient_name_stub = "_".join(recipient_name.split())
                recipients.append(recipient_name_stub)

            recipients_length = len(recipients)

            self.__check_pdf_length(recipients_length)

            print(
                f"Processed {recipients_length} recipient(s) for certificate ripping")

        return recipients

    def __check_pdf_length(self, recipients_length):
        pdf_length = self.pdf_length - (self.start_page_index)
        if pdf_length != recipients_length:
            raise ValueError(
                f"Number of recipients in CSV list ({recipients_length}) does not match with PDF length ({pdf_length})"
            )


if __name__ == "__main__":
    load_dotenv()

    ripper = CertRipper(
        mode=CertRipper.CSV_LIST_MODE,
        start_page_index=5,
        master_pdf_path=os.getenv("MASTER_PDF_PATH"),
        recipients_path=os.getenv("RECIPIENTS_PATH"),
        output_path=os.getenv("OUTPUT_PATH"),
        output_file_name_template=os.getenv("OUTPUT_FILE_NAME_TEMPLATE"),
    )

    ripper.process()
