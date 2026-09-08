"""
LinkedIn Saved Jobs Scraper
===========================
Automates the extraction of all saved jobs from your LinkedIn account
using Playwright for browser automation and exports to a formatted Excel file.

Usage:
    python scraper.py                  # Headless mode (default)
    python scraper.py --headful        # Visible browser (for debugging)
    python scraper.py --slow-mo 500    # Slow down actions by 500ms
"""

import argparse
import getpass
import os
import random
import re
import sys
import time
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from playwright.sync_api import TimeoutError as PlaywrightTimeout
from playwright.sync_api import sync_playwright

# ──────────────────────────────────────────────
# Constants
# ──────────────────────────────────────────────
LINKEDIN_LOGIN_URL = "https://www.linkedin.com/login"
LINKEDIN_SAVED_JOBS_URL = "https://www.linkedin.com/my-items/saved-jobs/"
TWO_FA_WAIT_TIMEOUT = 120_000  # 120 seconds for manual 2FA
SCROLL_PAUSE_MIN = 2.0
SCROLL_PAUSE_MAX = 4.0
ACTION_DELAY_MIN = 1.5
ACTION_DELAY_MAX = 3.5
MAX_SCROLL_ATTEMPTS = 100  # safety limit


def human_delay(low=ACTION_DELAY_MIN, high=ACTION_DELAY_MAX):
    """Sleep for a randomised duration to mimic human behaviour."""
    time.sleep(random.uniform(low, high))


# ──────────────────────────────────────────────
# 1. Credential Loading
# ──────────────────────────────────────────────
def load_credentials():
    """Load LinkedIn credentials from .env file or prompt interactively."""
    load_dotenv()
    email = os.getenv("LINKEDIN_EMAIL")
    password = os.getenv("LINKEDIN_PASSWORD")

    if not email:
        email = input("LinkedIn email: ").strip()
    if not password:
        password = getpass.getpass("LinkedIn password: ")

    if not email or not password:
        print("❌  Email and password are required.")
        sys.exit(1)

    return email, password


# ──────────────────────────────────────────────
# 2. Login
# ──────────────────────────────────────────────
def _find_and_fill(page, selectors, value, field_name, timeout=30_000):
    """Try multiple selectors to find an input field, then fill it."""
    for selector in selectors:
        try:
            page.wait_for_selector(selector, state="visible", timeout=timeout)
            page.fill(selector, value)
            print(f"   ✔ Found {field_name} field: {selector}")
            return True
        except PlaywrightTimeout:
            continue
        except Exception:
            continue
    print(f"   ❌ Could not find {field_name} field with any selector.")
    return False


def login(page, email, password):
    """Log in to LinkedIn, handling potential 2FA / captcha challenges."""
    print("🔑  Navigating to LinkedIn login …")
    page.goto(LINKEDIN_LOGIN_URL, wait_until="networkidle", timeout=60_000)
    human_delay(2, 4)

    # Wait a bit more for JS to render
    page.wait_for_load_state("domcontentloaded")

    # Try multiple selectors for the username field
    username_selectors = [
        'input#username',
        'input[name="session_key"]',
        'input[autocomplete="username"]',
        'input[type="text"]',
    ]
    if not _find_and_fill(page, username_selectors, email, "email/username"):
        print("❌  Cannot find the email input field. LinkedIn may have changed their page.")
        print("    Try running with --headful to see what the page looks like.")
        sys.exit(1)

    human_delay(0.5, 1.2)

    # Try multiple selectors for the password field
    password_selectors = [
        'input#password',
        'input[name="session_password"]',
        'input[autocomplete="current-password"]',
        'input[type="password"]',
    ]
    if not _find_and_fill(page, password_selectors, password, "password"):
        print("❌  Cannot find the password input field. LinkedIn may have changed their page.")
        print("    Try running with --headful to see what the page looks like.")
        sys.exit(1)

    human_delay(0.5, 1.0)

    # Try multiple selectors for the submit button
    submit_selectors = [
        'button[type="submit"]',
        'button[data-litms-control-urn="login-submit"]',
        'button.btn__primary--large',
    ]
    for selector in submit_selectors:
        try:
            btn = page.query_selector(selector)
            if btn:
                btn.click()
                print(f"   ✔ Clicked submit: {selector}")
                break
        except Exception:
            continue

    # Wait for navigation — either feed or a verification challenge
    print("⏳  Waiting for login to complete …")
    try:
        page.wait_for_url(
            re.compile(r"(feed|my-items|checkpoint|challenge)"),
            timeout=30_000,
        )
    except PlaywrightTimeout:
        pass  # Will check below

    # Check for 2FA / verification challenge
    current_url = page.url
    if any(k in current_url for k in ("checkpoint", "challenge", "two-step")):
        print(
            "\n🔐  Two-factor / verification challenge detected!\n"
            "    Please complete the verification in the browser window.\n"
            f"    Waiting up to {TWO_FA_WAIT_TIMEOUT // 1000}s …\n"
        )
        try:
            page.wait_for_url(
                re.compile(r"(feed|my-items)"),
                timeout=TWO_FA_WAIT_TIMEOUT,
            )
        except PlaywrightTimeout:
            print("❌  Timed out waiting for 2FA. Exiting.")
            sys.exit(1)

    # Final check — are we actually logged in?
    if "login" in page.url:
        print("❌  Login failed. Check your credentials.")
        sys.exit(1)

    print("✅  Logged in successfully!")
    human_delay()


# ──────────────────────────────────────────────
# 3. Scroll & Collect Job Cards
# ──────────────────────────────────────────────
def scroll_and_collect_jobs(page):
    """Scroll through the saved-jobs page to load all cards, then return them."""
    print("📂  Navigating to Saved Jobs …")
    page.goto(LINKEDIN_SAVED_JOBS_URL, wait_until="domcontentloaded")
    human_delay(3, 5)

    # Wait for the job list container to appear
    try:
        page.wait_for_selector(
            ".scaffold-layout__list, .jobs-search-results-list, .reusable-search__entity-result-list",
            timeout=15_000,
        )
    except PlaywrightTimeout:
        # Try alternate selectors for different page layouts
        try:
            page.wait_for_selector("ul.scaffold-layout__list-container", timeout=10_000)
        except PlaywrightTimeout:
            print("⚠️  Could not detect the saved jobs list. The page layout may have changed.")

    prev_count = 0
    no_change_rounds = 0

    for attempt in range(1, MAX_SCROLL_ATTEMPTS + 1):
        # Scroll to bottom
        page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        pause = random.uniform(SCROLL_PAUSE_MIN, SCROLL_PAUSE_MAX)
        time.sleep(pause)

        # Count visible job cards
        job_cards = page.query_selector_all(
            "li.reusable-search__result-container, "
            "li.scaffold-layout__list-item, "
            "li.jobs-search-results__list-item, "
            "div.job-card-container, "
            "li.ember-view.jobs-search-results__list-item"
        )
        current_count = len(job_cards)

        if current_count > prev_count:
            print(f"   ↻ Scroll #{attempt}: loaded {current_count} jobs …")
            prev_count = current_count
            no_change_rounds = 0
        else:
            no_change_rounds += 1

        # If count hasn't changed for 3 consecutive scrolls, we've reached the end
        if no_change_rounds >= 3:
            break

    print(f"✅  Found {prev_count} saved jobs total.")
    return page.query_selector_all(
        "li.reusable-search__result-container, "
        "li.scaffold-layout__list-item, "
        "li.jobs-search-results__list-item, "
        "div.job-card-container, "
        "li.ember-view.jobs-search-results__list-item"
    )


# ──────────────────────────────────────────────
# 4. Extract Job Details
# ──────────────────────────────────────────────
def safe_text(element):
    """Return trimmed inner text or 'N/A' if element is None."""
    if element is None:
        return "N/A"
    text = element.inner_text().strip()
    return text if text else "N/A"


def safe_attr(element, attr):
    """Return an attribute value or 'N/A'."""
    if element is None:
        return "N/A"
    val = element.get_attribute(attr)
    return val.strip() if val else "N/A"


def extract_job_details(page, card, index, total):
    """Click into a job card and scrape all available detail fields."""
    job = {
        "title": "N/A",
        "company": "N/A",
        "job_url": "N/A",
        "company_url": "N/A",
        "location": "N/A",
        "employment_type": "N/A",
        "seniority_level": "N/A",
        "description": "N/A",
        "date_posted": "N/A",
        "industry": "N/A",
        "job_function": "N/A",
    }

    try:
        # ---- Grab basics from the card itself ----
        link_el = card.query_selector("a.job-card-container__link, a.job-card-list__title, a[href*='/jobs/view/']")
        if link_el:
            job["title"] = safe_text(link_el)
            href = safe_attr(link_el, "href")
            if href != "N/A":
                # Normalise to full URL
                if href.startswith("/"):
                    href = "https://www.linkedin.com" + href
                # Strip tracking params
                job["job_url"] = href.split("?")[0]

        company_el = card.query_selector(
            ".job-card-container__primary-description, "
            ".artdeco-entity-lockup__subtitle, "
            "a.job-card-container__company-name, "
            "span.job-card-container__primary-description"
        )
        job["company"] = safe_text(company_el)

        location_el = card.query_selector(
            ".job-card-container__metadata-item, "
            ".artdeco-entity-lockup__caption, "
            "span.job-card-container__metadata-wrapper"
        )
        job["location"] = safe_text(location_el)

        date_el = card.query_selector("time, .job-card-container__listed-time, .job-card-container__footer-item")
        job["date_posted"] = safe_text(date_el)

        # ---- Click into detail panel ----
        clickable = link_el or card
        try:
            clickable.click()
            human_delay(1.5, 3.0)
        except Exception:
            print(f"   ⚠️  Could not click job #{index + 1}, using card-only data.")
            return job

        # Wait for detail pane to load
        try:
            page.wait_for_selector(
                ".jobs-unified-top-card, .job-details-jobs-unified-top-card__job-title, "
                ".jobs-details__main-content, .jobs-description",
                timeout=8_000,
            )
        except PlaywrightTimeout:
            pass

        # ---- Title from detail panel (more reliable) ----
        detail_title = page.query_selector(
            ".job-details-jobs-unified-top-card__job-title, "
            ".jobs-unified-top-card__job-title, "
            "h1.t-24, h2.jobs-unified-top-card__job-title"
        )
        if detail_title:
            t = safe_text(detail_title)
            if t != "N/A":
                job["title"] = t

        # ---- Company from detail panel ----
        detail_company = page.query_selector(
            ".job-details-jobs-unified-top-card__company-name a, "
            ".jobs-unified-top-card__company-name a, "
            "a.jobs-unified-top-card__company-name"
        )
        if detail_company:
            c = safe_text(detail_company)
            if c != "N/A":
                job["company"] = c
            company_href = safe_attr(detail_company, "href")
            if company_href != "N/A":
                if company_href.startswith("/"):
                    company_href = "https://www.linkedin.com" + company_href
                job["company_url"] = company_href.split("?")[0]

        # ---- Location from detail panel ----
        detail_location = page.query_selector(
            ".job-details-jobs-unified-top-card__bullet, "
            ".jobs-unified-top-card__bullet, "
            ".jobs-unified-top-card__workplace-type"
        )
        if detail_location:
            loc = safe_text(detail_location)
            if loc != "N/A":
                job["location"] = loc

        # ---- Criteria items (seniority, employment type, industry, function) ----
        criteria_items = page.query_selector_all(
            ".job-details-jobs-unified-top-card__job-insight span, "
            ".jobs-unified-top-card__job-insight span, "
            "li.jobs-description-details__list-item, "
            ".jobs-description__job-criteria-item"
        )
        for item in criteria_items:
            text = safe_text(item).lower()
            # Employment type
            if any(k in text for k in ("full-time", "part-time", "contract", "temporary",
                                        "internship", "volunteer", "freelance")):
                job["employment_type"] = safe_text(item)
            # Seniority
            elif any(k in text for k in ("entry", "associate", "mid-senior", "director",
                                          "executive", "senior", "intern")):
                job["seniority_level"] = safe_text(item)
            # Industry
            elif any(k in text for k in ("industry", "industries")):
                job["industry"] = safe_text(item)
            # Job function
            elif any(k in text for k in ("function", "functions")):
                job["job_function"] = safe_text(item)

        # Try structured criteria (header + value pairs)
        criteria_headers = page.query_selector_all(
            ".jobs-description__job-criteria-subheader, "
            ".job-details-jobs-unified-top-card__job-insight-view-model-secondary"
        )
        criteria_values = page.query_selector_all(
            ".jobs-description__job-criteria-text, "
            ".job-details-jobs-unified-top-card__job-insight-view-model-secondary + span"
        )
        for header, value in zip(criteria_headers, criteria_values):
            h = safe_text(header).lower()
            v = safe_text(value)
            if "seniority" in h:
                job["seniority_level"] = v
            elif "employment" in h or "type" in h:
                job["employment_type"] = v
            elif "industry" in h:
                job["industry"] = v
            elif "function" in h:
                job["job_function"] = v

        # ---- Description ----
        desc_el = page.query_selector(
            ".jobs-description__content, "
            ".jobs-description-content__text, "
            ".jobs-box__html-content, "
            "#job-details"
        )
        if desc_el:
            d = safe_text(desc_el)
            if d != "N/A":
                # Trim excessive whitespace but keep paragraph structure
                job["description"] = re.sub(r'\n{3,}', '\n\n', d).strip()

        # ---- Date posted from detail panel ----
        date_detail = page.query_selector(
            ".jobs-unified-top-card__posted-date, "
            ".job-details-jobs-unified-top-card__primary-description-container span, "
            "span.tvm__text--neutral"
        )
        if date_detail:
            dt = safe_text(date_detail)
            if dt != "N/A" and any(k in dt.lower() for k in ("ago", "hour", "day", "week", "month", "just")):
                job["date_posted"] = dt

    except Exception as exc:
        print(f"   ⚠️  Error extracting job #{index + 1}: {exc}")

    print(f"   ✔ [{index + 1}/{total}]  {job['title'][:55]}  —  {job['company'][:30]}")
    return job


# ──────────────────────────────────────────────
# 5. Excel Export
# ──────────────────────────────────────────────
COLUMNS = [
    ("Job Title", "title"),
    ("Company", "company"),
    ("Job URL", "job_url"),
    ("Company URL", "company_url"),
    ("Location", "location"),
    ("Employment Type", "employment_type"),
    ("Seniority Level", "seniority_level"),
    ("Date Posted", "date_posted"),
    ("Industry", "industry"),
    ("Job Function", "job_function"),
    ("Description", "description"),
]


def export_to_excel(jobs, output_dir):
    """Write scraped jobs to a formatted Excel workbook."""
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    filename = output_dir / f"saved_jobs_{timestamp}.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "Saved Jobs"

    # ---- Styles ----
    header_font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color="0A66C2", end_color="0A66C2", fill_type="solid")  # LinkedIn blue
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell_alignment = Alignment(vertical="top", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    link_font = Font(name="Calibri", color="0563C1", underline="single", size=10)
    normal_font = Font(name="Calibri", size=10)
    alt_fill = PatternFill(start_color="F2F7FC", end_color="F2F7FC", fill_type="solid")

    # ---- Headers ----
    for col_idx, (header, _) in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border

    # ---- Data Rows ----
    for row_idx, job in enumerate(jobs, start=2):
        is_alt = (row_idx % 2 == 0)
        for col_idx, (_, key) in enumerate(COLUMNS, start=1):
            value = job.get(key, "N/A")
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.border = thin_border
            cell.alignment = cell_alignment
            cell.font = normal_font
            if is_alt:
                cell.fill = alt_fill

            # Make URLs clickable hyperlinks
            if key in ("job_url", "company_url") and value not in ("N/A", ""):
                cell.value = value
                cell.hyperlink = value
                cell.font = link_font
            else:
                cell.value = value

    # ---- Auto-fit column widths ----
    for col_idx, (header, key) in enumerate(COLUMNS, start=1):
        if key == "description":
            ws.column_dimensions[get_column_letter(col_idx)].width = 60
        elif key in ("job_url", "company_url"):
            ws.column_dimensions[get_column_letter(col_idx)].width = 45
        elif key == "title":
            ws.column_dimensions[get_column_letter(col_idx)].width = 40
        elif key == "company":
            ws.column_dimensions[get_column_letter(col_idx)].width = 30
        elif key == "location":
            ws.column_dimensions[get_column_letter(col_idx)].width = 25
        else:
            ws.column_dimensions[get_column_letter(col_idx)].width = 18

    # ---- Freeze top row & auto-filter ----
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    # ---- Row height for description column ----
    for row in range(2, len(jobs) + 2):
        ws.row_dimensions[row].height = 60

    wb.save(filename)
    return filename


# ──────────────────────────────────────────────
# 6. Main
# ──────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description="Scrape saved jobs from LinkedIn")
    parser.add_argument("--headful", action="store_true", help="Run browser in visible (headed) mode")
    parser.add_argument("--slow-mo", type=int, default=0, help="Slow down Playwright actions by N ms")
    parser.add_argument("--output", type=str, default="output", help="Output directory for Excel file")
    args = parser.parse_args()

    email, password = load_credentials()

    print("\n" + "=" * 60)
    print("   LinkedIn Saved Jobs Scraper")
    print("=" * 60 + "\n")

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=not args.headful,
            slow_mo=args.slow_mo,
        )
        context = browser.new_context(
            viewport={"width": 1280, "height": 900},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            ),
            locale="en-US",
        )
        page = context.new_page()

        # Step 1 — Login
        login(page, email, password)

        # Step 2 — Scroll & collect
        job_cards = scroll_and_collect_jobs(page)

        if not job_cards:
            print("❌  No saved jobs found. The page layout may have changed.")
            browser.close()
            sys.exit(1)

        # Step 3 — Extract details
        total = len(job_cards)
        print(f"\n📋  Extracting details from {total} saved jobs …\n")
        jobs = []
        for i, card in enumerate(job_cards):
            job_data = extract_job_details(page, card, i, total)
            jobs.append(job_data)
            human_delay(1.0, 2.5)

        browser.close()

    # Step 4 — Export
    print(f"\n💾  Exporting {len(jobs)} jobs to Excel …")
    output_path = export_to_excel(jobs, args.output)
    print(f"✅  Done!  File saved to: {output_path.resolve()}\n")


if __name__ == "__main__":
    main()
