# Cert Ripper

A Python tool that can separate certificate PDF files of recipients from a single PDF file. It also uses `pywin` to automate Outlook email drafting and sending.

## Usage

1. Install `pywin` using `pip`.
2. Make sure you have Outlook with remembered credentials setup on your desktop.
3. Provide the necessary files for input.
   - Recipients CSV - CSV file containing recipients full name and email
   - Template HTML - the actual HTML-formatted message to be sent
   - Master PDF - the single PDF that contains individual certificates per page
4. Provide `.env`. Copy from `.env.example`. Tweak accordingly.
5. Run `ripper.py` first to acquired extracted PDFs of recipients from master PDF.
6. Run `mailer.py` if you want to prepare Outlook email drafts that will be sent to the recipients.
   - I recommend just saving drafts so you have a chance to verify the emails before sending them.

## To-do

1. Implement batch sending so the script can directly send the emails. Outlook limits 30 outbound messages per minute. Use some thread sleeping operation.
2. Unit tests...
