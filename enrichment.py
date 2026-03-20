import requests
from bs4 import BeautifulSoup
from datetime import datetime
import openpyxl
import re
import time
import os

FILE = "startups.xlsx"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )
}

# Column indices (1-based)
COL_COMPANY       = 1
COL_FOUNDER_NAME  = 6
COL_FOUNDER_EMAIL = 7
COL_WEBSITE       = 8
COL_COMPANY_EMAIL = 14   # new column


# ── UTILITIES ─────────────────────────────────────────────────────────────────

def clean_company_name(name):
    """Strip descriptor prefixes like 'Edtech startup Lenskart' -> 'Lenskart'"""
    prefixes = (
        r"^(?:e[\s-]?commerce|ecommerce|edtech|fintech|healthtech|health tech|"
        r"d2c|foodtech|food tech|gaming|travel|fashion|beauty|wellness|"
        r"fitness|media|retail|consumer|mobility|insurtech|logistics|"
        r"payments|lending|materials science|music licensing|legal[\s-]?tech|"
        r"ai|robotics|rural commerce|urban|delhi[\s-]?based|mumbai[\s-]?based|"
        r"bangalore[\s-]?based|bengaluru[\s-]?based|india[\s-]?us|"
        r"solar|ev|electric vehicle|battery|healthcare|mortgage|biotech|"
        r"longevity|science|platform|app|brand|accelerator|indian fabless|fabless)"
        r"(?:\s+(?:startup|company|firm|platform|maker|brand|based))?\s+"
    )
    cleaned = re.sub(prefixes, '', name, flags=re.IGNORECASE).strip().strip("'\"").strip()
    return cleaned if cleaned else name


def name_to_domain_guesses(name):
    """Turn a company name into a list of likely domain names to try."""
    clean = name.lower()
    clean = re.sub(r"['\-\s]+", "", clean)   # remove spaces, hyphens, apostrophes
    clean = re.sub(r"[^a-z0-9]", "", clean)  # remove any other special chars

    # Also try with spaces replaced by nothing and common TLDs
    tlds = [".com", ".in", ".co.in", ".io", ".co"]
    guesses = []
    for tld in tlds:
        guesses.append(f"https://www.{clean}{tld}")
        guesses.append(f"https://{clean}{tld}")
    return guesses


def fetch_page(url, timeout=10):
    try:
        resp = requests.get(url, headers=HEADERS, timeout=timeout, allow_redirects=True)
        if resp.status_code == 200 and len(resp.text) > 500:
            return resp.text
    except Exception:
        pass
    return None


# ── WEBSITE FINDER ────────────────────────────────────────────────────────────

def find_website(company_name):
    """Try common domain patterns directly — no Google needed."""
    clean_name = clean_company_name(company_name)
    guesses = name_to_domain_guesses(clean_name)

    for url in guesses:
        html = fetch_page(url)
        if html:
            # Confirm it's actually the right company
            soup = BeautifulSoup(html, "html.parser")
            page_text = soup.get_text(separator=" ").lower()
            # Check if company name words appear on the page
            words = [w for w in clean_name.lower().split() if len(w) > 3]
            if any(w in page_text for w in words):
                # Return clean base URL
                from urllib.parse import urlparse
                parsed = urlparse(url)
                return f"{parsed.scheme}://{parsed.netloc}"
    return ""


# ── EMAIL EXTRACTION ──────────────────────────────────────────────────────────

PERSONAL_BLACKLIST = [
    "noreply", "no-reply", "support", "info", "hello", "contact",
    "admin", "team", "press", "media", "legal", "privacy",
    "careers", "jobs", "help", "sales", "marketing", "newsletter",
    "feedback", "enquiry", "inquiry", "invoice", "billing",
]

def extract_emails(html):
    text = BeautifulSoup(html, "html.parser").get_text(separator=" ")
    pattern = r'[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}'
    all_emails = list(set(re.findall(pattern, text)))

    personal = [e for e in all_emails if not any(b in e.lower() for b in PERSONAL_BLACKLIST)]
    company  = [e for e in all_emails if any(b in e.lower() for b in ["contact", "info", "hello", "team"])]

    return personal, company


def scrape_emails_from_site(base_url):
    """
    Visit homepage + common contact/about pages.
    Returns (founder_email, company_email).
    """
    pages_to_try = [
        base_url,
        base_url.rstrip("/") + "/about",
        base_url.rstrip("/") + "/about-us",
        base_url.rstrip("/") + "/team",
        base_url.rstrip("/") + "/our-team",
        base_url.rstrip("/") + "/contact",
        base_url.rstrip("/") + "/contact-us",
        base_url.rstrip("/") + "/founders",
    ]

    all_personal = []
    all_company  = []

    for url in pages_to_try:
        html = fetch_page(url, timeout=8)
        if not html:
            continue
        personal, company = extract_emails(html)
        all_personal.extend(personal)
        all_company.extend(company)
        time.sleep(0.5)

    # Deduplicate
    all_personal = list(dict.fromkeys(all_personal))
    all_company  = list(dict.fromkeys(all_company))

    founder_email = all_personal[0] if all_personal else ""
    company_email = all_company[0] if all_company else (all_personal[1] if len(all_personal) > 1 else "")

    return founder_email, company_email


# ── FOUNDER NAME EXTRACTION ───────────────────────────────────────────────────

def extract_founder_name(html):
    soup = BeautifulSoup(html, "html.parser")
    text = soup.get_text(separator=" ")

    patterns = [
        r'(?:founded by|co-founded by)\s+([A-Z][a-z]+(?:\s+[A-Z][a-z]+){1,2})',
        r'([A-Z][a-z]+(?:\s+[A-Z][a-z]+){1,2})\s*[,–\-]\s*(?:Founder|Co-Founder|CEO|MD|Managing Director)',
        r'(?:Founder|Co-Founder|CEO|MD|Managing Director)\s*[,–\-]\s*([A-Z][a-z]+(?:\s+[A-Z][a-z]+){1,2})',
    ]
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            name = match.group(1).strip()
            if 2 <= len(name.split()) <= 4:
                return name
    return ""


def find_founder_name(base_url):
    pages_to_try = [
        base_url.rstrip("/") + "/about",
        base_url.rstrip("/") + "/about-us",
        base_url.rstrip("/") + "/team",
        base_url.rstrip("/") + "/our-team",
        base_url.rstrip("/") + "/founders",
        base_url,
    ]
    for url in pages_to_try:
        html = fetch_page(url, timeout=8)
        if html:
            name = extract_founder_name(html)
            if name:
                return name
        time.sleep(0.5)
    return ""


# ── MAIN ENRICHMENT ───────────────────────────────────────────────────────────

def enrich_row(company_name, existing_website=""):
    result = {
        "website": existing_website,
        "founder_name": "",
        "founder_email": "",
        "company_email": "",
    }

    clean_name = clean_company_name(company_name)
    print(f"  [{clean_name}]")

    # Step 1: Find website
    website = existing_website
    if not website:
        website = find_website(clean_name)
        result["website"] = website

    if not website:
        print(f"    Website: not found — skipping email search")
        return result

    print(f"    Website: {website}")

    # Step 2: Find founder name
    founder_name = find_founder_name(website)
    result["founder_name"] = founder_name
    print(f"    Founder: {founder_name or 'not found'}")

    # Step 3: Scrape emails
    founder_email, company_email = scrape_emails_from_site(website)
    result["founder_email"] = founder_email
    result["company_email"] = company_email
    print(f"    Founder email:  {founder_email or 'not found'}")
    print(f"    Company email:  {company_email or 'not found'}")

    return result


def run_enrichment():
    print("\n=== ENRICHMENT STARTED ===")
    print(f"Time: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n")

    if not os.path.exists(FILE):
        print(f"ERROR: {FILE} not found. Run scraper.py first.")
        return

    wb = openpyxl.load_workbook(FILE)
    ws = wb.active

    updated = 0
    skipped = 0

    for row in ws.iter_rows(min_row=2):
        company       = row[COL_COMPANY - 1].value
        founder_name  = row[COL_FOUNDER_NAME - 1].value
        founder_email = row[COL_FOUNDER_EMAIL - 1].value
        website       = row[COL_WEBSITE - 1].value
        company_email = row[COL_COMPANY_EMAIL - 1].value

        # Skip rows already fully enriched
        if founder_email and company_email:
            skipped += 1
            continue

        if not company:
            continue

        data = enrich_row(
            str(company).strip(),
            existing_website=str(website).strip() if website else ""
        )

        # Write back whatever we found
        if data["website"] and not website:
            row[COL_WEBSITE - 1].value = data["website"]
        if data["founder_name"] and not founder_name:
            row[COL_FOUNDER_NAME - 1].value = data["founder_name"]
        if data["founder_email"] and not founder_email:
            row[COL_FOUNDER_EMAIL - 1].value = data["founder_email"]
        if data["company_email"] and not company_email:
            row[COL_COMPANY_EMAIL - 1].value = data["company_email"]

        if any([data["founder_name"], data["founder_email"], data["company_email"]]):
            updated += 1

        wb.save(FILE)
        time.sleep(2)

    print(f"\nEnrichment complete.")
    print(f"  Rows updated: {updated}")
    print(f"  Rows skipped (already done): {skipped}")
    print("=== ENRICHMENT DONE ===\n")


if __name__ == "__main__":
    run_enrichment()