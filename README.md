# Invoice-Monitor

![Python](https://img.shields.io/badge/python-3.10%2B-blue?logo=python&logoColor=white)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey?logo=windows)
![License](https://img.shields.io/badge/license-MIT-green)
![Tests](https://github.com/Ajeje97/Invoice-Monitor/actions/workflows/tests.yml/badge.svg)

A lightweight Python automation tool that monitors a Microsoft Outlook inbox for incoming invoices (Nota Fiscal / NF-e), creates task alerts, and logs all detections to a CSV file for audit purposes

 Table of Contents

Motivation
How It Works
Requirements
Installation
Usage
Output
Project Structure
Design Decisions
Known Limitations
Roadmap
License
Author


Motivation
In manufacturing environments, delayed invoice processing causes payment failures, financial penalties, and damaged relationships with suppliers and internal teams. Manual monitoring of a shared inbox is error-prone — emails are missed, especially during high-volume periods.
This tool automates the detection of invoice-related emails in Microsoft Outlook, creates a task alert with a due date and reminder, and maintains a traceable CSV log of every detection. It was designed to reduce human error in a real operational context where missing an invoice has measurable business consequences.

How It Works
Outlook Inbox
     │
     ▼
[Filter by date & read status]
     │
     ▼
[Keyword scan: subject → body → attachments]
     │
     ├─ NF detected ──► Create Outlook Task (due: next day, reminder: 5 min)
     │                          │
     │                          ▼
     │                   Log to CSV (audit trail)
     │
     └─ Not detected ──► Skip
Detection checks three layers in order of priority:

Subject line — keywords like nota fiscal, nfe, nf-e, danfe
Email body — same keywords, limited to first 5,000 characters for performance
Attachment filenames — regex patterns matching common NF file naming conventions


Requirements

Windows OS
Microsoft Outlook installed and configured with a valid account
Python 3.10 or higher
pywin32 library


Installation
1. Clone the repository
bashgit clone https://github.com/Ajeje97/invoice-monitor.git
cd invoice-monitor
2. Install dependencies
bashpip install pywin32
3. Add MIT License to your repository (if not yet added)
On GitHub: go to your repository → Add file → Create new file → name it LICENSE → click Choose a license template → select MIT → fill in your name → commit.

Usage
Run with default settings (last 7 days, unread emails only)
bashpython TarefaNF.py
Customize the scan window and recipient
bashpython TarefaNF.py --dias 14 --responsavel "Gabriel" --limite-alertas 30
Include already-read emails
bashpython TarefaNF.py --inclui-lidos
Save log to a custom path
bashpython TarefaNF.py --csv C:\relatorios\alertas_nf.csv
Run the built-in self-test (no Outlook required)
bashpython TarefaNF.py --autoteste
All available arguments
ArgumentDefaultDescription--dias7How many days back to scan--inclui-lidosFalseInclude already-read emails--responsavel"Equipe Fiscal"Responsible person named in the task--limite-alertas20Max tasks created per run--csvalertas_nf.csvOutput CSV file path--autotesteFalseRun offline self-test without Outlook

Output
Terminal
3 email(s) detected. 3 task(s) created.
Log saved to: C:\Users\...\alertas_nf.csv
Outlook Task — created automatically for each detected email:

Subject: Dar entrada no fiscal | [original email subject]
Due date: next business day
Reminder: 5 minutes after script execution

CSV log (alertas_nf.csv) — appended on every run:
data_registroremetenteassuntorecebido_emmotivo_deteccaotarefa_criadaerro2025-03-10 09:00:00Fornecedor AEnvio de NF-e pedido 1232025-03-10 08:45assuntosim

Project Structure
invoice-monitor/
│
├── TarefaNF.py          # Main script
├── alertas_nf.csv       # Auto-generated log (gitignored)
├── README.md
└── LICENSE

Note: Add alertas_nf.csv to your .gitignore to avoid committing potentially sensitive invoice data to version control.


Design Decisions
Why check subject → body → attachments in that order?
Subject-line matches are the fastest and most reliable signal. Scanning the full body is slower and more prone to false positives, so it only runs if the subject check fails. Attachment inspection is last because it requires iterating COM objects, which is the most expensive operation.
Why was the generic "nota" keyword removed?
Early testing showed it caused significant false positives — meeting notes, internal memos, and document names all matched. Only specific compound terms (nota fiscal, nf-e, etc.) are used to keep precision high.
Why limit body inspection to 5,000 characters?
Outlook emails can contain large HTML bodies, quoted reply chains, and signatures. Scanning the full body of every email would be slow. The first 5,000 characters reliably capture the relevant content of a genuine invoice notification.
Why use getattr with defaults instead of direct attribute access?
The Outlook COM interface can raise exceptions for certain email types (calendar invites, corrupted items, encrypted messages). Defensive attribute access via getattr(email, "Subject", "") prevents the entire scan from failing due to a single problematic item.
Why append to CSV instead of overwrite?
The CSV functions as an audit trail across multiple runs. Overwriting would destroy historical detection records, which are important for accountability in financial workflows.

Known Limitations

Windows only — depends on pywin32 and the Outlook COM interface; not compatible with macOS or Linux
Default inbox only — subfolders and shared mailboxes are not scanned in the current version
No deduplication — if the same email is detected across multiple runs (e.g., using --inclui-lidos), it will be logged and a new task created each time
No execution health log — the CSV only records detections, not whether the script ran successfully; a failed run produces no output
Requires Outlook to be running — the COM connection depends on an active Outlook session


Roadmap

 Migrate self-test to pytest for standard test reporting
 Add --demo mode: generates a sample CSV with fictional data without requiring Outlook
 Add execution log (timestamp + success/failure) separate from the detections CSV
 Support scanning subfolders and shared mailboxes
 Deduplication logic to avoid repeat task creation for the same email
 .gitignore template included in the repository


Automating with Windows Task Scheduler
To run this script automatically every few hours during business hours:

Open Task Scheduler → Create Basic Task
Set trigger: Daily, repeat every 2 hours between 08:00–18:00
Set action: Start a program

Program: python
Arguments: C:\path\to\TarefaNF.py --responsavel "YourName"


Check "Run only when user is logged on" (required for Outlook COM access)


License
This project is licensed under the MIT License.

Author
Ajeje97 — github.com/Ajeje97