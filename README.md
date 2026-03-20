# India B2C Startup Outreach Agent — v1

An automated agent that scrapes the internet daily for Indian B2C startups that have raised Series A or Series B funding, builds a database in Excel, finds founder contact details, and sends personalised cold emails — tracking replies automatically.

---

## What It Does

1. **Scrapes** funding news from Entrackr, Inc42, and Google News RSS every day
2. **Filters** for B2C startups in India that raised Series A or Series B in the last 12 months
3. **Enriches** each startup with founder name, founder email, and company email by visiting their website
4. **Sends** personalised cold emails via your Gmail account
5. **Tracks** email status (Not Sent / Sent – No Reply / Replied) in Excel, updated daily

---

## Project Structure

```
startup_agent/
│
├── setup_excel.py      # Run once — creates the Excel database with all columns
├── scraper.py          # Module 1 — finds funded startups from news sources
├── enrichment.py       # Module 2 — finds founder name, founder email, company email
├── email_sender.py     # Module 3 — sends personalised cold emails via Gmail
├── reply_tracker.py    # Module 4 — checks Gmail inbox for replies and updates Excel
├── main.py             # Orchestrator — runs all 4 modules in sequence
└── startups.xlsx       # The Excel database (auto-created, do not rename)
```

---

## Excel Database Schema

The file `startups.xlsx` has 14 columns:

| # | Column | Description |
|---|--------|-------------|
| 1 | Company Name | Startup name extracted from news |
| 2 | Sector | e.g. Consumer Tech, Fintech, Edtech |
| 3 | Funding Round | Series A or Series B |
| 4 | Amount Raised | e.g. $10M or Rs. 50 Cr |
| 5 | Date Funded | Date the funding was announced |
| 6 | Founder Name | Extracted from company website |
| 7 | Founder Email | Personal email found on website |
| 8 | Website | Official company website URL |
| 9 | Source URL | News article where funding was found |
| 10 | Email Status | `Not sent` / `Sent – no reply` / `Replied` |
| 11 | Email Sent Date | Timestamp when outreach email was sent |
| 12 | Reply Date | Timestamp when founder replied |
| 13 | Notes | Manual field for your own notes |
| 14 | Company Email | Generic contact email (info@, contact@) |

---

## Setup Instructions

### Step 1 — Install Python

Download Python from [python.org](https://www.python.org/downloads/).

During installation, check both boxes:
- ✅ Add python.exe to PATH
- ✅ Use admin privileges when installing

### Step 2 — Fix PATH (Windows only)

If you have other Python installations (e.g. from msys64), make sure the new Python appears first in your Environment Variables PATH.

Verify with:
```bash
where python
python -m pip --version
```

The first result of `where python` should point to your AppData Python, not msys64.

### Step 3 — Install dependencies

```bash
python -m pip install playwright beautifulsoup4 openpyxl requests googlesearch-python lxml
python -m playwright install chromium
```

### Step 4 — Set up Gmail App Password

The agent sends emails via your Gmail using an App Password (not your real password).

1. Go to [myaccount.google.com](https://myaccount.google.com)
2. Search for **"App Passwords"**
3. Create a new App Password for "Mail" on "Windows Computer"
4. Copy the 16-character password — you'll need it in `email_sender.py`

### Step 5 — Create the Excel database

```bash
cd C:\Users\harsh\startup_agent
python setup_excel.py
```

This creates `startups.xlsx` with all 14 columns and proper formatting.

---

## Running the Agent

### Run everything at once (recommended)

```bash
python main.py
```

This runs all 4 modules in sequence: scraper → enrichment → reply tracker → email sender.

### Run individual modules

```bash
python scraper.py        # Find new funded startups
python enrichment.py     # Find founder emails and websites
python reply_tracker.py  # Check Gmail for replies
python email_sender.py   # Send emails to new contacts
```

### Full reset (start from scratch)

```bash
del startups.xlsx
python setup_excel.py
python scraper.py
python enrichment.py
```

---

## Setting Up Daily Automation (Windows Task Scheduler)

To run the agent automatically every morning:

1. Press **Windows key**, search **Task Scheduler**, open it
2. Click **Create Basic Task** in the right panel
3. Name it: `Startup Outreach Agent`
4. Trigger: **Daily** at **8:00 AM**
5. Action: **Start a program**
   - Program: `C:\Users\harsh\AppData\Local\Python\bin\python.exe`
   - Arguments: `main.py`
   - Start in: `C:\Users\harsh\startup_agent`
6. Click **Finish**

The agent will now run every morning automatically as long as your laptop is on.

---

## How Each Module Works

### Module 1 — Scraper (`scraper.py`)

Scrapes three sources daily:

- **Entrackr** — reads funding keywords directly from article URL slugs
- **Inc42** — parses article headlines from the `/buzz/` and `/startups/` pages
- **Google News RSS** — queries Google's free RSS feed for Series A/B India startup news

Filters applied:
- Skips B2B companies (enterprise, SaaS, procurement, etc.)
- Skips roundup articles ("From X to Y", "Weekly Funding Report", etc.)
- Only saves rows where the company name is clean (no slashes, no URLs, max 6 words)
- Deduplicates by company name — never adds the same startup twice

### Module 2 — Enrichment (`enrichment.py`)

For each startup in the Excel sheet that is missing contact info:

1. **Finds the website** by guessing domain names directly:
   - Tries `companyname.com`, `companyname.in`, `companyname.co.in`, `companyname.io`
   - Confirms it's the right site by checking if the company name appears on the page
2. **Finds the founder name** by scraping `/about`, `/team`, `/founders` pages
   - Looks for patterns like "Founded by John Smith" or "Jane Doe — CEO"
3. **Finds emails** by scraping all contact pages
   - Founder Email: personal-looking emails (e.g. `rahul@company.com`)
   - Company Email: generic emails (e.g. `contact@company.com`, `info@company.com`)

Saves progress to Excel after every single row — safe to interrupt and resume.

**Expected hit rate for v1:**
- Website found: ~50–60% of startups
- Founder name: ~30–40%
- At least one email: ~30–40%

### Module 3 — Email Sender (`email_sender.py`)

For each row where:
- Email Status is `Not sent`
- At least one email address exists (founder or company)

The agent sends a personalised cold email using your Gmail via SMTP. The template uses the founder's name, company name, and funding round from the Excel row. After sending:
- Email Status → `Sent – no reply`
- Email Sent Date → current timestamp

Uses Gmail App Password for authentication (never stores your real password).

### Module 4 — Reply Tracker (`reply_tracker.py`)

Connects to your Gmail inbox via IMAP each morning and checks for replies from any email address already in the Excel sheet. If a reply is found:
- Email Status → `Replied`
- Reply Date → date of reply

---

## Data Sources

| Source | Type | Cost |
|--------|------|------|
| Entrackr | Web scraping | Free |
| Inc42 | Web scraping | Free |
| Google News RSS | RSS feed | Free |
| Company websites | Direct scraping | Free |

No paid APIs required for v1.

---

## Known Limitations (v1)

- **Enrichment hit rate is ~30–40%** — many startups don't publicly list founder emails. Rows with missing emails can be filled manually.
- **Company names occasionally have descriptor prefixes** — e.g. "Edtech startup Lenskart" instead of "Lenskart". These can be cleaned manually in Excel.
- **Requires laptop to be on** for the daily scheduler to trigger. If you want always-on automation, consider moving to a cloud server (AWS/GCP, ~$5/month).
- **YourStory is not scraped** — it blocks automated requests. Google News RSS covers similar stories.
- **LinkedIn is not used** — LinkedIn aggressively blocks scrapers. Founder names come from company websites only.

---

## Troubleshooting

**`pip` not recognised**
```bash
python -m pip install ...   # use python -m pip instead of pip
```

**Wrong Python being used (msys64)**
Fix PATH in Environment Variables — move `C:\Users\harsh\AppData\Local\Python\bin` above `C:\msys64\ucrt64\bin`.

**Scraper finds 0 results**
Run the diagnostic to check which sites are reachable:
```bash
python -c "
import requests
headers = {'User-Agent': 'Mozilla/5.0'}
for url in ['https://entrackr.com', 'https://inc42.com/buzz/']:
    r = requests.get(url, headers=headers, timeout=15)
    print(r.status_code, len(r.text), url)
"
```

**Enrichment crashes with IndexError**
The Excel file is missing column 14. Run:
```bash
python -c "
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
wb = openpyxl.load_workbook('startups.xlsx')
ws = wb.active
cell = ws.cell(row=1, column=14, value='Company Email')
cell.font = Font(bold=True, color='FFFFFF', size=11)
cell.fill = PatternFill(start_color='4B0082', end_color='4B0082', fill_type='solid')
cell.alignment = Alignment(horizontal='center', vertical='center')
ws.column_dimensions['N'].width = 30
wb.save('startups.xlsx')
print('Fixed.')
"
```

**Gmail authentication fails**
Make sure you are using an App Password, not your Gmail login password. App Passwords are 16 characters with no spaces. 2-Factor Authentication must be enabled on your Google account first.

---

## Roadmap (v2 ideas)

- Add Apollo.io or Hunter.io API for higher email hit rates
- Add LinkedIn scraping via a proxy service
- Move to a cloud server for always-on daily runs
- Add a Slack or WhatsApp notification when a founder replies
- Build a simple web dashboard to view and manage the pipeline
