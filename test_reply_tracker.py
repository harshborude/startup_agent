import imaplib, email, re, openpyxl, os
from email.header import decode_header
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()

FILE           = "test_startups.xlsx"   # points to test DB
GMAIL_ADDRESS  = os.getenv("GMAIL_ADDRESS")
GMAIL_APP_PASS = os.getenv("GMAIL_APP_PASS")

COL_FOUNDER_EMAIL = 7
COL_COMPANY_EMAIL = 14
COL_EMAIL_STATUS  = 10
COL_REPLY_DATE    = 12

def get_sent_addresses(ws):
    sent = {}
    for row in ws.iter_rows(min_row=2):
        status        = row[COL_EMAIL_STATUS - 1].value
        founder_email = row[COL_FOUNDER_EMAIL - 1].value or ""
        company_email = row[COL_COMPANY_EMAIL - 1].value or ""
        if status == "Sent – no reply":
            addr = founder_email or company_email
            if addr:
                sent[addr.lower().strip()] = row
    return sent

def decode_str(s):
    if s is None:
        return ""
    parts = decode_header(s)
    result = ""
    for part, encoding in parts:
        if isinstance(part, bytes):
            result += part.decode(encoding or "utf-8", errors="replace")
        else:
            result += str(part)
    return result

def fetch_reply_senders(mail):
    senders = set()
    try:
        mail.select("INBOX")
        _, message_ids = mail.search(None, "ALL")
        ids = message_ids[0].split()
        recent_ids = ids[-200:] if len(ids) > 200 else ids
        for msg_id in recent_ids:
            _, msg_data = mail.fetch(msg_id, "(RFC822)")
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg         = email.message_from_bytes(response_part[1])
                    from_header = decode_str(msg.get("From", ""))
                    match       = re.search(r'[\w.\+\-]+@[\w.\-]+\.\w+', from_header)
                    if match:
                        senders.add(match.group(0).lower().strip())
    except Exception as e:
        print(f"  Error scanning inbox: {e}")
    return senders

def run_test_reply_tracker():
    print("\n=== TEST REPLY TRACKER STARTED ===")
    print(f"Time: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print(f"Database: {FILE}\n")

    if not GMAIL_APP_PASS or GMAIL_APP_PASS == "YOUR_APP_PASSWORD_HERE":
        print("ERROR: GMAIL_APP_PASS not set in .env")
        return

    if not os.path.exists(FILE):
        print(f"ERROR: {FILE} not found. Run: python create_test_db.py")
        return

    wb = openpyxl.load_workbook(FILE)
    ws = wb.active

    sent_addresses = get_sent_addresses(ws)
    if not sent_addresses:
        print("No 'Sent – no reply' rows in test DB yet.")
        print("Run test_email_sender.py first, then re-run this.")
        print("=== TEST REPLY TRACKER DONE ===\n")
        return

    print(f"Tracking replies for: {list(sent_addresses.keys())}")

    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com", 993)
        mail.login(GMAIL_ADDRESS, GMAIL_APP_PASS)
        print("Connected to Gmail.\n")
    except imaplib.IMAP4.error:
        print("ERROR: IMAP auth failed.")
        print("Enable IMAP: Gmail → Settings → Forwarding and POP/IMAP → Enable IMAP")
        return
    except Exception as e:
        print(f"ERROR: {e}")
        return

    print("Scanning inbox for replies...")
    reply_senders = fetch_reply_senders(mail)
    mail.logout()
    print(f"Found {len(reply_senders)} unique senders in inbox.\n")

    # Show which tracked addresses were found in inbox
    print("Cross-referencing:")
    updated = 0
    for address, row in sent_addresses.items():
        company = row[0].value
        if address in reply_senders:
            print(f"  MATCH — {company} <{address}> → marking as Replied")
            row[COL_EMAIL_STATUS - 1].value = "Replied"
            row[COL_REPLY_DATE - 1].value   = datetime.now().strftime("%Y-%m-%d")
            updated += 1
        else:
            print(f"  NO REPLY — {company} <{address}>")

    if updated:
        wb.save(FILE)
        print(f"\n{updated} row(s) updated to 'Replied'.")
    else:
        print("\nNo replies found yet.")
        print("To simulate a reply: send an email from harshborude11@gmail.com to itself,")
        print("then run this script again.")

    print("=== TEST REPLY TRACKER DONE ===\n")

if __name__ == "__main__":
    run_test_reply_tracker()