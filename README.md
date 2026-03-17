# 📦 Carrier Email Automation System

## Overview
This project automates the processing of carrier arrival notice emails by:

- Extracting shipment identifiers (MBL / Container No.)
- Matching against a shipment database
- Automatically forwarding emails to responsible operators
- Preventing duplicate processing using logging

---

## 🚀 Key Features

- Outlook email automation (win32com)
- PDF text extraction (pdfplumber)
- Regex-based shipment detection
- Intelligent routing logic
- Duplicate email prevention
- Logging & audit trail

---

## 🧠 How It Works

1. Monitor Outlook folder for incoming emails
2. Read subject, body, and PDF attachments
3. Extract shipment identifiers using regex
4. Match against shipment lookup table
5. Forward email to correct operator(s)
6. Log processed emails to prevent duplication

---

## 🛠 Tech Stack

- Python
- Outlook COM API
- PDF processing
- Regular Expressions
- Data processing (pandas)

---

## 📁 Project Structure

carrier-email-automation/
│
├── main.py
├── requirements.txt
├── logs/
├── data/
└── README.md


---

## ⚠️ Note

This is a **sanitized portfolio version**:
- No real company data
- No internal email addresses
- No production environment details

---

## 💡 Use Case

Designed for logistics operations teams to reduce manual email handling and improve processing speed.

---

## 👨‍💻 Author

Jimmy Lin
