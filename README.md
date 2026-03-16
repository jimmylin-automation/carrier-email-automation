\# Carrier Email Automation System



Python automation tool that processes carrier arrival notice emails and automatically routes them to the correct shipment operators.



\## Diagram:



Carrier Email

&#x20;    ↓

Automation Script

&#x20;    ↓

Extract shipment numbers

&#x20;    ↓

Match with shipment database

&#x20;    ↓

Forward to responsible operator





\## Features



\- Outlook email automation

\- PDF attachment parsing

\- Shipment number extraction

\- Excel shipment lookup

\- Automatic operator routing

\- Duplicate prevention

\- Logging system



\## Technology



\- Python

\- pandas

\- pdfplumber

\- win32com

\- regex



\## Example Workflow



1\. Carrier sends arrival notice

2\. Script scans email subject/body/PDF

3\. Shipment number extracted

4\. Operator identified via Excel

5\. Email automatically forwarded



\## Purpose



Designed to reduce manual email sorting in freight forwarding operations.

