# LinkedIn Jobs Autosave & Scraper

An automated bot that **searches for jobs based on your criteria, automatically saves them**, and exports all details (including full job descriptions) into a set formatted Excel spreadsheet for further LLM-based job fit analysis.

## Features

- 🔍 **Job Autosaving**: Runs searches for multiple keywords, loops through pages, and automatically clicks "Save" on matching jobs.
- 🇩🇪 **German Ad Skipper**: Intelligently analyzes text to detect and skip job ads written in German.
- ⏭️ **Deduplication**: Automatically skips jobs you have already "Applied" to, "Archived", or are "Interviewing" for.
- 📊 **Excel Export**: Aggregates all saved jobs into a single `.xlsx` file.
- 🤖 **Human-like execution**: Uses randomized delays and Playwright to avoid anti-bot detection.

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

## Setup

### 1. Install dependencies
```bash
pip install -r requirements.txt
playwright install chromium
```

### 2. Configure criteria
Create a `.env` file in the same directory as the script:

```env
LINKEDIN_EMAIL=your_real_email@example.com
LINKEDIN_PASSWORD=your_real_password

# Autosave Search Configuration
SEARCH_KEYWORDS=Psychology, Customer Experience, Human Resources
SEARCH_LOCATION=Germany
SEARCH_TIME=86400  # 86400=24h, 604800=1week, 2592000=1month
AUTOSAVE_MAX_PAGES=5
```

## Usage

### Visible browser (Recommended to monitor behavior)
```bash
python scraper.py --headful
```

### Custom output directory
```bash
python scraper.py --headful --output my_jobs
```

## Disclaimer
Automated scraping may violate LinkedIn's terms of service. Use responsibly and at your own risk.
