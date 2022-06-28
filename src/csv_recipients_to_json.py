from dotenv import load_dotenv
import csv
import json
import os


class CSVRecipientsToJSON:

    def __init__(self, csv_recipients_path: str = None, json_recipients_path: str = None):
        self.csv_recipients_path = csv_recipients_path
        self.json_recipients_path = json_recipients_path

    def process(self) -> None:
        recipients = []

        with open(self.csv_recipients_path) as csv_file:
            csv_reader = csv.reader(csv_file, delimiter=",")

            for row in csv_reader:
                recipient_name = row[0]
                recipient_email = row[1]
                recipients.append({
                    "name": recipient_name,
                    "email": recipient_email
                })

        with open(self.json_recipients_path, "w") as out:
            json.dump(recipients, out)

        recipients_length = len(recipients)

        print(
            f"Converted \x1b[95m{recipients_length} recipient(s)\x1b[0m from CSV to JSON. \x1b[92mOutputted in path: {self.json_recipients_path}\x1b[0m")


if __name__ == "__main__":
    load_dotenv()

    converter = CSVRecipientsToJSON(
        csv_recipients_path=os.getenv("CSV_RECIPIENTS_PATH"),
        json_recipients_path=os.getenv("JSON_RECIPIENTS_PATH")
    )

    converter.process()
