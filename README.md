# LinkedIn Saved Jobs Scraper

Automatically extract **all** your saved jobs from LinkedIn into a beautifully formatted Excel spreadsheet.

## What It Extracts

| Field | Example |
|---|---|
| Job Title | Senior Software Engineer |
| Company | Google |
| Job URL | https://www.linkedin.com/jobs/view/123456 |
| Company URL | https://www.linkedin.com/company/google |
| Location | Amsterdam, Netherlands (Hybrid) |
| Employment Type | Full-time |
| Seniority Level | Mid-Senior level |
| Date Posted | 2 weeks ago |
| Industry | Technology |
| Job Function | Engineering |
| Description | Full job description text |

## Prerequisites

- **Python 3.9+** — [Download](https://www.python.org/downloads/)
- A LinkedIn account with saved jobs

## Setup

### 1. Install Python dependencies

```bash
cd linkedin-saved-jobs-scraper
pip install -r requirements.txt
```

### 2. Install Playwright browsers

```bash
playwright install chromium
```

### 3. Configure credentials

Copy the example env file and fill in your LinkedIn credentials:

```bash
copy .env.example .env
```

Edit `.env`:
```
LINKEDIN_EMAIL=your_real_email@example.com
LINKEDIN_PASSWORD=your_real_password
```

> **Tip:** If you prefer not to store your password in a file, leave it blank and the script will prompt you interactively.

## Usage

### Basic (headless)
```bash
python scraper.py
```

### Visible browser (recommended for first run)
```bash
python scraper.py --headful
```

### Custom output directory
```bash
python scraper.py --output my_jobs
```

### Slow-motion mode (debugging)
```bash
python scraper.py --headful --slow-mo 500
```

## Output

The script generates an Excel file in the `output/` folder (or your custom directory):

```
output/saved_jobs_2026-08-06_123456.xlsx
```

Features of the Excel file:
- 📊 LinkedIn-blue headers with bold white text
- 🔗 Clickable hyperlinks for job and company URLs
- 🎨 Alternating row colours for readability
- 📌 Frozen header row with auto-filter enabled
- 📏 Auto-sized columns

## Two-Factor Authentication

If your LinkedIn account has 2FA enabled:

1. Run with `--headful` so you can see the browser
2. The script will **pause automatically** when it detects a verification challenge
3. Complete the verification manually in the browser window
4. The script resumes automatically after verification (up to 120s timeout)

## Troubleshooting

| Problem | Solution |
|---|---|
| Login fails | Double-check credentials in `.env`. Try `--headful` to see what's happening |
| No jobs found | LinkedIn may have changed their page layout. Open an issue |
| CAPTCHA appears | Use `--headful` and solve it manually; the script will wait |
| Script is slow | This is intentional — randomised delays help avoid detection |
| Account warning | Stop using the script and wait 24h. Consider reducing frequency of use |

## ⚠️ Disclaimer

This tool is for **personal use only**. Automated access to LinkedIn may violate their [Terms of Service](https://www.linkedin.com/legal/user-agreement). Use at your own risk. The author is not responsible for any account restrictions.
