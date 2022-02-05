# Cert Ripper

A Python tool that can separate certificate PDF files of recipients from a single PDF file. It uses `pywin32` to automate Outlook email drafting and sending.

## Dependencies

Assuming you have Python and PIP installed, run:
```
$ pip install python-dotenv pdfplumber PyPDF2 pywin32
```

## Usage

1. Install Outlook for Windows. Make sure you have it with remembered credentials setup on your desktop.
2. Provide the necessary files for input.
   - Recipients CSV - CSV file containing recipients full name and email
   - Template HTML - the actual HTML-formatted message to be sent
   - Master PDF - the single PDF that contains individual certificates per page
3. Provide `.env`. Copy from `.env.example`. Tweak accordingly.
4. Run `ripper.py` first to acquire extracted PDFs of recipients from master PDF.
5. Run `mailer.py` if you want to prepare Outlook email drafts that will be sent to the recipients.
   - I recommend just saving drafts so you have a chance to verify the emails before sending them.

## To-do

1. Implement batch sending so the script can directly send the emails. Outlook limits 30 outbound messages per minute. Use some thread sleeping operation.
2. Unit tests...
