# 📊 Daily Operations Reporting Automation

This project automates the daily disbursement reporting workflow for operations teams in banking and finance using Python.

## 🚀 Features
- Reads disbursement data from Excel
- Generates a clean, formatted daily report (Excel)
- Summarizes disbursed & rejected amounts
- Sends an automatic email with summary + attachment

## 🛠 Tech Stack
- Python
- Pandas
- OpenPyXL
- smtplib (email)
- Excel

## 📁 Project Structure
daily-ops-report-automation/
├── data/
│ └── disbursements_2025-07-26.xlsx # Raw input Excel file
├── output/
│ └── daily_report_2025-07-26.xlsx # Final formatted report
├── screenshots/
│ └── email_preview.png # Screenshot of received email
├── daily_ops_report.py # Main Python automation script
