import imaplib
import email
import re
import openpyxl
import os
from email.header import decode_header
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()

# ── CONFIG ────────────────────────────────────────────────────────────────────
GMAIL_ADDRESS  = os.getenv("GMAIL_ADDRESS")
GMAIL_APP_PASS = os.getenv("GMAIL_APP_PASS")
FILE           = "startups.xlsx"

# ── COLUMN INDICES (1-based) ──────────────────────────────────────────────────
COL_FOUNDER_EMAIL = 7
COL_COMPANY_EMAIL = 14
COL_EMAIL_STATUS  = 10
COL_REPLY_DATE    = 12


# ── HELPERS ───────────────────────────────────────────────────────────────────

def get_sent_addresses(ws):
    """Return {email: row} for all rows with status 'Sent – no reply'."""
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
    """Return a set of all sender email addresses found in the inbox."""
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


# ── MAIN ──────────────────────────────────────────────────────────────────────

def run_reply_tracker():
    print("\n=== REPLY TRACKER STARTED ===")
    print(f"Time: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n")

    if not GMAIL_APP_PASS or GMAIL_APP_PASS == "YOUR_APP_PASSWORD_HERE":
        print("ERROR: GMAIL_APP_PASS not set in .env file.")
        return

    if not os.path.exists(FILE):
        print(f"ERROR: {FILE} not found.")
        return

    wb = openpyxl.load_workbook(FILE)
    ws = wb.active

    sent_addresses = get_sent_addresses(ws)
    if not sent_addresses:
        print("No 'Sent – no reply' rows found. Nothing to track.")
        print("=== REPLY TRACKER DONE ===\n")
        return

    print(f"Tracking replies for {len(sent_addresses)} sent emails...")

    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com", 993)
        mail.login(GMAIL_ADDRESS, GMAIL_APP_PASS)
        print("Connected to Gmail.\n")
    except imaplib.IMAP4.error:
        print("ERROR: IMAP authentication failed.")
        print("Enable IMAP in Gmail: Settings → See all settings → Forwarding and POP/IMAP → Enable IMAP")
        return
    except Exception as e:
        print(f"ERROR: {e}")
        return

    print("Scanning inbox...")
    reply_senders = fetch_reply_senders(mail)
    mail.logout()
    print(f"Found {len(reply_senders)} unique senders.\n")

    updated = 0
    for address, row in sent_addresses.items():
        if address in reply_senders:
            company = row[0].value
            print(f"  REPLY: {company} <{address}>")
            row[COL_EMAIL_STATUS - 1].value = "Replied"
            row[COL_REPLY_DATE - 1].value   = datetime.now().strftime("%Y-%m-%d")
            updated += 1

    if updated:
        wb.save(FILE)
        print(f"\n{updated} row(s) updated to 'Replied'.")
    else:
        print("No new replies found.")

    print("=== REPLY TRACKER DONE ===\n")


if __name__ == "__main__":
    run_reply_tracker()