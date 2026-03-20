from scraper import run_scraper
from enrichment import run_enrichment
from reply_tracker import run_reply_tracker
from email_sender import run_email_sender
from datetime import datetime

def run_agent():
    print("=" * 50)
    print("   STARTUP OUTREACH AGENT — DAILY RUN")
    print(f"   {datetime.now().strftime('%A, %d %B %Y — %H:%M')}")
    print("=" * 50)

    # Step 1: Find new funded startups
    run_scraper()

    # Step 2: Enrich with founder details and emails
    run_enrichment()

    # Step 3: Check inbox for replies before sending new emails
    run_reply_tracker()

    # Step 4: Send emails to new contacts
    run_email_sender()

    print("=" * 50)
    print("   ALL DONE. Check startups.xlsx for updates.")
    print("=" * 50)

if __name__ == "__main__":
    run_agent()