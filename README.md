# Cert Ripper

A Python tool that can separate certificate PDF files of recipients from a single PDF file. It uses `pywin32` to automate Outlook email drafting and sending.

## Dependencies

## Usage

### Setting Up

After cloning this repository and assuming you have Python and PIP installed, run:

```
$ pip install -r requirements.txt
```

You need to have [Outlook for Windows](https://www.office.com) installed. Set it up and login to your Outlook account.

### Environment variables

Create a `.env` file and copy it from `.env.example`. You need to configure the following variables:

| Variable                | Description                                                                         | Parameters                                                                                              |
| ----------------------- | ----------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------- |
| `CC_EMAIL`              | The CC emails for each certificate (separated by `;`)                               |                                                                                                         |
| `SUBJECT`               | The e-mail subject for each certificate                                             | `certificate` - the name of the certificate group                                                       |
| `START_PAGE_INDEX`      | The page that the ripper should start in (starts at `0`). Useful for skipping pages |                                                                                                         |
| `RIPPED_CERTS_PATH`     | The folder where the separated certificates are placed                              |                                                                                                         |
| `RIPPED_CERT_FILE_NAME` | The file name of ripped certificates                                                | `index` - page number, `tag` - short name of certificate group, `recipient` - slugged name of recipient |
| `MASTER_PDF_PATH`       | The file location where the master PDF is found                                     |
| `EMAIL_TEMPLATE_PATH`   | The file location where the template HTML is found                                  |
| `CSV_RECIPIENTS_PATH`   | The file location where the list of recipients in CSV is found                      |
| `JSON_RECIPIENTS_PATH`  | The file location where the list of recipients in JSON is found                     |
| `JSON_POINTS_PATH`      | The file location where the JSON certificate points are found                       |

### Creating inputs

You need to have the following files prepared to start ripping and mailing certificates:

- Master PDF - the single PDF that contains individual certificates per page
- Points JSON - the description of how the certificates are structured for ripping and mailing in the master PDF
- Template HTML - the HTML-formatted template that is used as the body for each created e-mail

#### Converting CSV Recipients

You can convert a list of CSV recipients to a JSON array of recipient objects by using the `csv_recipients_to_json.py` script.

Each recipient object has the following fields:

| Field   | Description                       |
| ------- | --------------------------------- |
| `name`  | _string_. Name of the recipient   |
| `email` | _string_. E-mail of the recipient |

You can then use the recipient objects in structuring the points JSON array.

#### Points

Points are defined as a JSON array below like so:

```json
[
  {
    "name": "Special Award Certificate",
    "tag": "SA",
    "recipients": [
      {
        "name": "Juan Dela Cruz",
        "email": "jdc@some.org"
      }
    ]
  },
  {
    "name": "Certificate of Participation",
    "tag": "CP",
    "recipients": [
      {
        "name": "Pedro Gil",
        "email": "pg@some.org"
      },
      {
        "name": "Maria Clara",
        "email": "mc@some.org"
      }
    ]
  }
]
```

Each object in the points array is called a point. Each point is considered as a certificate group. Likewise, `points[0]` is a group of special award certificates and `points[1]` is a group of participation certificates.

Each point has the following fields:

| Field        | Description                                                                                           |
| ------------ | ----------------------------------------------------------------------------------------------------- |
| `name`       | _string_. Name of the certificate group; may be used for email subject and body                       |
| `tag`        | _string_. Short name of the certificate group; may be used for ripped certificate file name           |
| `recipients` | array of `recipient` objects. A list of recipients who intends to receive a certificate of this group |

The points are a series of certificates. Juan's special award certificate is found at page 1 (unless specified to have a different `START_PAGE_INDEX`) in the master PDF. Pedro's participation certificate is found at page 2. Maria's at page 3 and so on.

### Running Scripts

- Run `ripper.py` first to acquire extracted PDFs of recipients from master PDF.
- Run `mailer.py` if you want to prepare Outlook email drafts that will be sent to the recipients.
  - I recommend just saving drafts so you have a chance to verify the emails before sending them.

## To-do

1. Implement batch sending so the script can directly send the emails. Outlook limits 30 outbound messages per minute. Use some thread sleeping operation.
2. Unit tests...
