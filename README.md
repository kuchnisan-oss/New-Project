# Outlook Inbox Contact Extractor (Microsoft Graph)

A Windows 11 desktop application that connects to Microsoft Graph, scans emails from a selected mailbox, extracts contact details, removes duplicates, and saves the data to a CSV file.

## Features
- Tkinter desktop UI for Tenant ID, Client ID, and mailbox email.
- Device Code authentication via MSAL.
- Fetches 500 newest emails per run and stores a cursor for the next batch.
- Extracts sender name, sender email, company (from domain), and a phone number when present.
- Deduplicates contacts by email before appending to CSV.
- Stores the CSV and cursor on the Windows desktop.

## Requirements
- Python 3.10+
- An Azure AD app registration with delegated **Mail.Read** permission.

Install dependencies:

```bash
pip install -r requirements.txt
```

## Run

```bash
python main.py
```

## Output
- CSV: `~/Desktop/outlook_contacts.csv`
- Cursor: `~/Desktop/outlook_contact_cursor.json`

Use the **Reset Cursor** button to start from the newest messages again.
