<h1 align="center">Merchant Report and Email Tool</h1>

<p align="center"><b>Generate merchant transaction reports and send them by email or WhatsApp</b></p>

<p align="center">
  <img src="https://img.shields.io/badge/Python-3.12-3776AB?logo=python&logoColor=white" alt="Python">
  <img src="https://img.shields.io/badge/platform-Windows-0078D6?logo=windows&logoColor=white" alt="Platform">
  <img src="https://img.shields.io/badge/updated-2026--04--18-success" alt="Updated">
  <a href="https://github.com/alisadhk/emil-report-mm/releases/latest"><img src="https://img.shields.io/badge/downloads-Releases-blue" alt="Releases"></a>
</p>

---

## Overview
A PyQt6 desktop application that loads transaction data files, builds a report for each merchant and sends it by email or through WhatsApp Web.

## Features
- Load transaction data files and a merchant list
- Per-merchant report generation
- Send by email or WhatsApp Web
- Application log panel
- Reconciles settlement amounts with transaction logs and shows a Match column
- Batch processing: formatted PDF and Excel reports for hundreds of merchants
- Omnichannel delivery with Outlook and WhatsApp, including auto-paste of pre-filled message templates

## How it works
- Load the three input files; the merchant table is built from them.
- For each merchant the tool computes the totals, builds a formatted report and stores it in the reports folder.
- The **Send** column opens your mail client or WhatsApp Web, attaches the report and fills the message template.
- *Report Status* and *Send Status* show what is done and what is pending.

## Requirements
- Python 3.10+ with PyQt6 and the report libraries from requirements.txt.
- Your own input files (no sample data is published).
- Outlook or WhatsApp Web for sending.

## Screenshots
<p align="center"><img src="screenshots/overview.png" alt="Merchant Report and Email Tool screenshot" width="720"></p>
<p align="center"><img src="screenshots/screenshot-1.png" alt="Merchant Report and Email Tool screenshot" width="720"></p>

## Download and install
### Option A - Download
No prebuilt binary is published for this project. Use Option B below.

### Option B - Run from source
Requirements: Windows 10/11 and Python 3.10+.

```bash
git clone https://github.com/alisadhk/emil-report-mm.git
cd emil-report-mm
python -m pip install -r requirements.txt
python merchant_report_tool2.py
```

## How to use
1. Start the program with `python merchant_report_tool2.py`.
2. Under **Load Data Files** pick the **Send Info File**, the **Clients File** and the **Transactions File**.
3. The **Merchant List** fills in (see the merchant counter).
4. Press **Generate All Reports**; the *Report Status* column shows progress for every merchant.
5. Send each report by email or WhatsApp Web (*Send* / *Send Status* columns) and use **Open Reports Folder** to see the files.
6. Open **Log** to review what happened; use **Clear** to start over.

## Configuration and privacy
- All sensitive values from the original private version (passwords, tokens, license codes, user names, internal IP addresses, personal names and links) were replaced with `*****`.
- No customer data, reports or spreadsheets are included in this repository.

## Troubleshooting
- Merchant list stays empty: check that all three files are loaded and use the expected columns.
- Nothing sent: sign in to WhatsApp Web / Outlook first.
- The overview image shows a newer build with POS and B2B tabs; this repository contains the earlier script.

## Build the executable yourself
```bash
python -m pip install pyinstaller
python -m PyInstaller --noconfirm 
```
The executable is created in the `dist/` folder.

## Notes
- Use only data you are authorized to process.

---

<p align="center">Maintained by <a href="https://github.com/alisadhk">@alisadhk</a> - last update 2026-04-18</p>