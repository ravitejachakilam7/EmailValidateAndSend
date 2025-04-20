# 📧 Automated Resume Email Sender

This Python script automates the process of sending your resume to multiple email addresses from an Excel file. It validates each email address for proper format and domain validity before sending, and attaches your resume in each email.

---

## ✨ Features

- ✅ Validates email format using regex
- 📬 Checks domain MX records (can the domain receive emails?)
- 📎 Sends resume as a PDF attachment
- 📊 Reads from an Excel sheet (expects a column named `Email`)
- ⏱ Waits every 10 emails to avoid spam flags or throttling

---

## 📁 File Structure

You only need the following in your repo:
