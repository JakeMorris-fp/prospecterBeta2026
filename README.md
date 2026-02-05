
# Prospecting Manager (Streamlit App)

A lightweight app for prospecting financial professionals. Import contacts from Excel, track call outcomes (voicemail / yes / no / call back later), capture meeting dates & notes, and generate personalized email text with placeholders for name and meeting details.

## ✨ Features

- **Excel import/export**: Bring in your prospect list from `.xlsx` (or `.xls`) and export updates back to Excel.
- **Contact export**: One-click CSV export of just **Name, Phone, Email**.
- **Call outcomes**: Track `Voicemail`, `Yes` (with meeting date/time & notes), `No`, and `Call back later` (with date/time).
- **Email templates**: Define subject/body templates with placeholders like `{name}`, `{first_name}`, `{company}`, `{meeting_date}`, `{meeting_time}` and preview per contact.
- **Mailto links**: Open a pre-filled email in your default email client, or download a CSV of personalized emails for bulk sending.

## 🧰 Requirements

- Python 3.9+
- See `requirements.txt` for Python packages.

## 🚀 Quick Start

1. **Create & activate a virtual environment (recommended)**
   ```bash
   python -m venv .venv
   # Windows
   .venv\Scripts\activate
   # macOS/Linux
   source .venv/bin/activate
   ```
2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```
3. **Run the app**
   ```bash
   streamlit run app.py
   ```
4. **In the app**
   - Download the **Excel template** from the sidebar (optional) and populate your contacts, or upload your own Excel file with columns `Name`, `Phone`, `Email` (optional: `Company`, `Title`).
   - Edit statuses, meeting/callback date-times, and notes directly in the table.
   - Export the updated dataset to Excel and/or export just Name/Phone/Email (CSV).
   - Configure your email templates, preview them, and generate a CSV of personalized emails.

## 📝 Columns & Placeholders

**Required columns in Excel**
- `Name`
- `Phone`
- `Email`

**Optional columns**
- `Company`
- `Title`

**App-added columns**
- `Status` (one of: `Voicemail`, `Yes`, `No`, `Call back later`)
- `MeetingDateTime` (datetime)
- `CallbackDateTime` (datetime)
- `Notes`

**Email template placeholders**
- `{name}` – full name
- `{first_name}` – first token of the name
- `{company}` – company (if provided)
- `{meeting_datetime}` – raw datetime string
- `{meeting_date}` – formatted date, e.g., `January 15, 2026`
- `{meeting_time}` – formatted time, e.g., `3:30 PM`

> Tip: Use mailto links for individual emails, or the generated CSV to drive bulk sends from Outlook/Excel/Power Automate or your CRM.

## 🔐 Privacy & Compliance

This app handles personally identifiable information (PII). Ensure you store and share files in compliance with your organization's policies and any applicable regulations. Consider restricting access and encrypting files at rest.

## 🧪 Troubleshooting

- If you upload `.xls` files and see an error, ensure `xlrd` is installed and that the file is not corrupted.
- If date/time formatting looks odd after pasting from other tools, edit directly in the app to normalize.

---
If you need SMTP sending, CRM integration, or .ics calendar invites, open an issue or ask and we can extend the app.
