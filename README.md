# AttendanceTracker

## Libraries Used:
- openpyxl: Excel file handling
- smtplib: Email sending via SMTP
- email.mime.multipart, email.mime.text: Email content creation
- os: Operating system interaction
- time: Script delay

## Overview:
This script monitors student attendance, sending email alerts if below a set threshold.

## Setup:
1. **Excel Workbook:** Load 'attendance.xlsx'.
2. **Global Variables:**
   - `ATTENDANCE_FILE`: Excel file name
   - `ATTENDANCE_THRESHOLD`: Minimum attendance ratio
   - `SMTP_SERVER`, `SMTP_PORT`: SMTP configuration
   - `SENDER_EMAIL`, `SENDER_PASSWORD`: Sender's email credentials
   - `EMAIL_LIST`: List of email addresses

## Functions:
- `save_excel(workbook)`: Save updated Excel sheet.
- `track_attendance(worksheet)`: Calculate attendance ratio, send email alerts if below threshold.
- `send_email(to_email)`: Send email alert.

## Email Configuration:
- Gmail SMTP: `smtp.gmail.com`, port `587`.
- **Note:** Avoid hardcoding sensitive information; use environment variables or a configuration file.

## Main Loop:
- Infinite loop: `while True`.
- Continuous attendance tracking and email alerts.
- Wait 1 hour (`time.sleep(3600)`) between iterations.
