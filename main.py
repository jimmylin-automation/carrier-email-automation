import win32com.client
import pandas as pd
import re
import pdfplumber
import os
import tempfile
import logging
import uuid
from datetime import datetime

# ==============================
# SETTINGS (Example Configuration)
# ==============================
SIMULATION_MODE = False

EXCEL_PATH = r"data/shipment_lookup.xlsx"
MAIL_FOLDER_NAME = "Carrier_Notifications"

START_DATE = datetime(2026, 1, 1)
END_DATE = datetime(2026, 12, 31)

LOG_FILE = "logs/automation_log.txt"
PROCESSED_FILE = "logs/processed_messages.txt"

# ==============================
# LOGGING
# ==============================

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s - %(message)s"
)

# ==============================
# LOAD PROCESSED EMAIL IDS
# ==============================

if os.path.exists(PROCESSED_FILE):
    with open(PROCESSED_FILE, "r") as f:
        processed_ids = set(line.strip() for line in f)
else:
    processed_ids = set()

# ==============================
# OPERATOR EMAIL MAP (Example)
# ==============================

OPERATOR_EMAIL_MAP = {
    "OPERATOR_A": "operatorA@example.com",
    "OPERATOR_B": "operatorB@example.com",
    "OPERATOR_C": "operatorC@example.com"
}

# ==============================
# DATA NORMALIZATION
# ==============================

def normalize(value):
    if pd.isna(value):
        return ""

    v = str(value).upper().strip()
    v = v.replace("/", "")
    v = v.replace(" ", "")

    return v

# ==============================
# SAFE PDF TEXT EXTRACTION
# ==============================

def extract_pdf_text(attachment):

    try:
        unique_name = str(uuid.uuid4()) + ".pdf"
        temp_path = os.path.join(tempfile.gettempdir(), unique_name)

        data = attachment.PropertyAccessor.GetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x37010102"
        )

        with open(temp_path, "wb") as f:
            f.write(data)

        text = ""

        with pdfplumber.open(temp_path) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    text += t + " "

        try:
            os.remove(temp_path)
        except:
            pass

        return text.upper()

    except Exception as e:
        logging.info(f"PDF read error: {e}")
        return ""

# ==============================
# LOAD SHIPMENT LOOKUP DATA
# ==============================

df = pd.read_excel(EXCEL_PATH)

COL_OPERATOR = "Operator"
COL_SHIPMENT = "Shipment_ID"
COL_MBL = "Master_Bill"
COL_CONTAINER = "Container_Number"

mbl_map = {}
container_map = {}

for _, row in df.iterrows():

    operator = str(row[COL_OPERATOR]).strip()
    shipment_id = str(row[COL_SHIPMENT]).strip()

    if operator not in OPERATOR_EMAIL_MAP:
        continue

    email = OPERATOR_EMAIL_MAP[operator]

    mbl = normalize(row[COL_MBL])
    if mbl:
        mbl_map[mbl] = (email, shipment_id)

    containers = str(row[COL_CONTAINER]).split(",")

    for c in containers:
        c_norm = normalize(c)
        if c_norm:
            container_map[c_norm] = (email, shipment_id)

logging.info("Lookup tables loaded")

# ==============================
# OUTLOOK CONNECTION
# ==============================

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
inbox = namespace.GetDefaultFolder(6)

target_folder = next(f for f in inbox.Folders if f.Name == MAIL_FOLDER_NAME)

messages = target_folder.Items
messages.Sort("[ReceivedTime]", True)

# ==============================
# PATTERNS
# ==============================

MBL_PATTERN = r"[A-Z]{2,5}[A-Z/]{0,10}\d{3,}"
CONTAINER_PATTERN = r"[A-Z]{4}\s?\d{7}"

# ==============================
# MATCH FUNCTION
# ==============================

def collect_matches(content):

    results = set()

    content = re.sub(r"[^A-Z0-9/ ]", " ", content.upper())

    for m in re.findall(MBL_PATTERN, content):

        m_norm = normalize(m)

        if m_norm in mbl_map:
            results.add(mbl_map[m_norm])

    for c in re.findall(CONTAINER_PATTERN, content):

        c_norm = normalize(c)

        if c_norm in container_map:
            results.add(container_map[c_norm])

    return results

# ==============================
# SIMULATION MODE (for demo / portfolio)
# ==============================

if SIMULATION_MODE:
    print("Running in SIMULATION MODE...\n")

    sample_emails = [
        {
            "subject": "Arrival Notice MBL123456",
            "body": "Container ABCD1234567 arrived"
        },
        {
            "subject": "Notice: Shipment update",
            "body": "Container XYZU7654321 available for pickup"
        }
    ]

    for msg in sample_emails:
        content = msg["subject"] + " " + msg["body"]
        matches = collect_matches(content)

        print("EMAIL:", content)
        print("MATCH RESULT:", matches)
        print("-" * 50)
        print(f"→ Forward To: {emails}")
        print(f"→ Shipment IDs: {syyz_list}")

    print("Simulation completed.")
    exit()

# ==============================
# PROCESS EMAILS
# ==============================

for msg in messages:

    try:

        received = msg.ReceivedTime.replace(tzinfo=None)

        if not (START_DATE <= received <= END_DATE):
            continue

        if msg.EntryID in processed_ids:
            continue

        all_matches = set()

        content = (msg.Subject or "") + " " + (msg.Body or "")
        all_matches |= collect_matches(content)

        if msg.Attachments.Count > 0:

            for i in range(1, msg.Attachments.Count + 1):

                att = msg.Attachments.Item(i)

                if att.FileName.lower().endswith(".pdf"):

                    pdf_text = extract_pdf_text(att)

                    if pdf_text:
                        all_matches |= collect_matches(pdf_text)

        if not all_matches:
            continue

        emails = sorted({m[0] for m in all_matches})
        shipment_list = sorted({m[1] for m in all_matches})

        fwd = msg.Forward()

        fwd.To = ";".join(emails)
        fwd.Subject = f"{', '.join(shipment_list)} | {msg.Subject}"

        fwd.Send()

        with open(PROCESSED_FILE, "a") as f:
            f.write(msg.EntryID + "\n")

        processed_ids.add(msg.EntryID)

        logging.info(f"Forwarded → {shipment_list} → {emails}")

    except Exception as e:
        logging.info(f"Error processing email: {e}")

logging.info("Automation run completed")
print("Automation completed")
