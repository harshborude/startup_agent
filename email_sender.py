import smtplib
import ssl
import openpyxl
import os
import time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()

# ── CONFIG ────────────────────────────────────────────────────────────────────
GMAIL_ADDRESS  = os.getenv("GMAIL_ADDRESS")
GMAIL_APP_PASS = os.getenv("GMAIL_APP_PASS")
SENDER_NAME    = "Harsh"
DAILY_LIMIT    = 15
FILE           = "startups.xlsx"

# ── COLUMN INDICES (1-based) ──────────────────────────────────────────────────
COL_COMPANY       = 1
COL_ROUND         = 3
COL_FOUNDER_NAME  = 6
COL_FOUNDER_EMAIL = 7
COL_EMAIL_STATUS  = 10
COL_EMAIL_SENT    = 11
COL_COMPANY_EMAIL = 14


# ── EMAIL TEMPLATE ────────────────────────────────────────────────────────────

def build_email(founder_name, company_name, funding_round):
    greeting = f"Hi {founder_name.split()[0]}," if founder_name else "Hi there,"
    subject  = f"Helping {company_name} scale on Instagram"

    body = f"""{greeting}

Congratulations on your {funding_round} raise — that's a big milestone and well deserved.

I came across {company_name} and was genuinely impressed by what you're building. I run a marketing agency that helps D2C brands scale their Instagram presence — from content strategy and creatives to paid social and influencer tie-ups.

Post a Series A/B raise, most consumer brands hit the same wall: great product, growing team, but Instagram either plateaus or burns budget without clear ROI. We've helped several funded D2C brands fix exactly that.

I'd love to explore if there's a fit — even a 20-minute call to share what's been working for brands at your stage would be worthwhile.

Would you be open to a quick chat this week or next?

Best,
{SENDER_NAME}
{GMAIL_ADDRESS}"""

    return subject, body


# ── SEND EMAIL ────────────────────────────────────────────────────────────────

def send_email(to_address, subject, body):
    try:
        msg            = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"]    = f"{SENDER_NAME} <{GMAIL_ADDRESS}>"
        msg["To"]      = to_address
        msg.attach(MIMEText(body, "plain"))

        context = ssl.create_default_context()
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
            server.login(GMAIL_ADDRESS, GMAIL_APP_PASS)
            server.sendmail(GMAIL_ADDRESS, to_address, msg.as_string())
        return True

    except smtplib.SMTPAuthenticationError:
        print("  ERROR: Gmail authentication failed. Check your App Password in .env")
        return False
    except smtplib.SMTPRecipientsRefused:
        print(f"  ERROR: Address rejected: {to_address}")
        return False
    except Exception as e:
        print(f"  ERROR: {e}")
        return False


# ── MAIN ──────────────────────────────────────────────────────────────────────

def run_email_sender():
    print("\n=== EMAIL SENDER STARTED ===")
    print(f"Time: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n")

    if not GMAIL_APP_PASS or GMAIL_APP_PASS == "YOUR_APP_PASSWORD_HERE":
        print("ERROR: GMAIL_APP_PASS not set in .env file.")
        print("Open .env and replace YOUR_APP_PASSWORD_HERE with your 16-character App Password.")
        return

    if not os.path.exists(FILE):
        print(f"ERROR: {FILE} not found.")
        return

    wb   = openpyxl.load_workbook(FILE)
    ws   = wb.active
    sent = 0

    for row in ws.iter_rows(min_row=2):
        if sent >= DAILY_LIMIT:
            print(f"\nDaily limit of {DAILY_LIMIT} reached. Stopping.")
            break

        company       = row[COL_COMPANY - 1].value
        funding_round = row[COL_ROUND - 1].value or "Series A/B"
        founder_name  = row[COL_FOUNDER_NAME - 1].value or ""
        founder_email = row[COL_FOUNDER_EMAIL - 1].value or ""
        company_email = row[COL_COMPANY_EMAIL - 1].value or ""
        status        = row[COL_EMAIL_STATUS - 1].value

        if not company:
            continue
        if status != "Not sent":
            continue

        to_address = founder_email or company_email
        if not to_address:
            print(f"  SKIP {company} — no email found")
            continue

        subject, body = build_email(founder_name, str(company), str(funding_round))
        print(f"  Sending → {company} <{to_address}>")
        success = send_email(to_address, subject, body)

        if success:
            row[COL_EMAIL_STATUS - 1].value = "Sent – no reply"
            row[COL_EMAIL_SENT - 1].value   = datetime.now().strftime("%Y-%m-%d %H:%M")
            wb.save(FILE)
            sent += 1
            print(f"  ✓ Sent ({sent}/{DAILY_LIMIT})")
            time.sleep(10)
        else:
            row[COL_EMAIL_STATUS - 1].value = "Send failed"
            wb.save(FILE)

    print(f"\nTotal sent today: {sent}")
    print("=== EMAIL SENDER DONE ===\n")


if __name__ == "__main__":
    run_email_sender()