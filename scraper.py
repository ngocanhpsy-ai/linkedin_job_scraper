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


def is_german(text):
    """Detect if text is primarily German based on common stop words."""
    if not text:
        return False
    text = text.lower()
    german_words = [" und ", " für ", " mit ", " oder ", " zu ", " auf ", " ein ", " eine ", " der ", " die ", " das ", " wir ", " suchen ", " werden ", " nicht ", " m/w/d "]
    match_count = sum(1 for word in german_words if word in text)
    return match_count >= 3


# ──────────────────────────────────────────────
# 1. Credential Loading
# ──────────────────────────────────────────────
def load_credentials():
    """Load LinkedIn credentials and search config from .env file next to the script."""
    # Look for .env in the same folder as this script
    script_dir = Path(__file__).resolve().parent
    env_path = script_dir / ".env"

    print(f"📁  Looking for .env at: {env_path}")

    if env_path.exists():
        load_dotenv(dotenv_path=env_path, override=True)
        print("   ✔ .env file found and loaded.")
    else:
        print(f"   ⚠️  No .env file found.")

    email = os.environ.get("LINKEDIN_EMAIL")
    password = os.environ.get("LINKEDIN_PASSWORD")
    
    # Search Config
    search_keywords_raw = os.environ.get("SEARCH_KEYWORDS", "")
    search_keywords_list = [k.strip() for k in search_keywords_raw.split(",") if k.strip()]
    
    search_location = os.environ.get("SEARCH_LOCATION", "Worldwide")
    search_time = os.environ.get("SEARCH_TIME", "86400") # default 24h
    autosave_max_pages = int(os.environ.get("AUTOSAVE_MAX_PAGES", "5"))

    if not email or not password:
        print("\n   [!] Email or password missing.")
        print("       Please update your .env file with:")
        print("       LINKEDIN_EMAIL=your_email")
        print("       LINKEDIN_PASSWORD=your_password")
        print("       SEARCH_KEYWORDS=Psychology, Customer Experience")
        print("       SEARCH_LOCATION=Germany")
        print("       SEARCH_TIME=86400")
        print("       AUTOSAVE_MAX_PAGES=5")
        sys.exit(1)

    print(f"   ✔ Credentials loaded for: {email}")
    if search_keywords_list:
        print(f"   ✔ Search Config: {len(search_keywords_list)} keywords in '{search_location}' (Max pages: {autosave_max_pages})")
    
    return email, password, search_keywords_list, search_location, search_time, autosave_max_pages


# ──────────────────────────────────────────────
# 2. Login
# ──────────────────────────────────────────────
def _wait_for_any_input(page, timeout=30_000):
    """Wait until at least one visible input element appears on the page."""
    print("    ⏳ Waiting for login form to render …")
    try:
        page.wait_for_selector("input:visible", state="visible", timeout=timeout)
        print("    ✔ Login form detected.")
        return True
    except PlaywrightTimeout:
        return False


def _debug_page_inputs(page):
    """Print all input elements on the page for debugging."""
    try:
        inputs = page.locator("input").element_handles()
        print(f"\n   🔍 DEBUG — Found {len(inputs)} input element(s) on page:")
        visible_count = 0
        for el in inputs:
            is_vis = el.is_visible()
            _id = el.get_attribute("id") or ""
            _name = el.get_attribute("name") or ""
            _type = el.get_attribute("type") or ""
            _auto = el.get_attribute("autocomplete") or ""
            _placeholder = el.get_attribute("placeholder") or ""
            icon = "👁 " if is_vis else "🚫 "
            if is_vis:
                visible_count += 1
            print(f"      {icon} id='{_id}' name='{_name}' type='{_type}' autocomplete='{_auto}' placeholder='{_placeholder}'")
        print()
        return visible_count
    except Exception as e:
        print(f"   ⚠️ Could not debug inputs: {e}")
        return 0


def login(page, email, password):
    """Logs into LinkedIn."""
    print("\n🔑  Navigating to LinkedIn login …")
    
    # 1) Go to the login page
    try:
        page.goto(LINKEDIN_LOGIN_URL, wait_until="domcontentloaded", timeout=60_000)
    except PlaywrightTimeout:
        print("   ⚠️ Timeout waiting for login page to load. Will try to proceed anyway...")

    print("    (waiting for page to load …)")
    
    try:
        page.wait_for_load_state("domcontentloaded", timeout=30000)
    except Exception:
        pass
    print("    ✔ DOM loaded. Waiting for JS to render the form …")

    human_delay(1, 2)

    # Dump inputs for debugging visibility
    visible_inputs = _debug_page_inputs(page)

    # Try to fill username natively
    username_filled = False
    try:
        email_field = page.locator('input[type="email"], input[type="text"]').filter(visible=True).first
        email_field.fill(email)
        print("   ✔ Filled email via native Playwright.")
        username_filled = True
    except Exception as e:
        print(f"   ❌ Native email fill failed: {e}")

    if not username_filled:
        print("❌  Cannot find the email input field.")
        sys.exit(1)

    human_delay(0.5, 1.2)

    # Try to fill password natively
    password_filled = False
    try:
        pw_field = page.locator('input[type="password"]').filter(visible=True).first
        pw_field.fill(password)
        print("   ✔ Filled password via native Playwright.")
        password_filled = True
    except Exception as e:
        print(f"   ❌ Native password fill failed: {e}")

    if not password_filled:
        print("❌  Cannot find the password input field.")
        sys.exit(1)

    human_delay(0.5, 1.0)
    
    # Debug screenshot before submitting
    import os
    from pathlib import Path
    output_dir = Path(__file__).resolve().parent / "output"
    output_dir.mkdir(parents=True, exist_ok=True)
    page.screenshot(path=str(output_dir / "debug_before_submit.png"), full_page=True)

    human_delay(0.5, 1.0)

    # Click submit
    try:
        pw_field = page.locator('input[type="password"]').filter(visible=True).first
        pw_field.press("Enter")
        print("   ✔ Pressed Enter in password field to submit.")
    except Exception as e:
        print(f"   ⚠️ Could not press Enter: {e}")
        
    # As a secondary fallback, click any button that says Einloggen
    try:
        for btn in page.locator('button:has-text("Sign in"), button:has-text("Einloggen")').all():
            if btn.is_visible() and btn.is_enabled():
                btn.click()
                print("   ✔ Clicked fallback submit button.")
                break
    except Exception:
        pass


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
        print("❌  Login failed. Check your credentials or captcha.")
        
        # Dump failure info
        import os
        from pathlib import Path
        output_dir = Path(__file__).resolve().parent / "output"
        output_dir.mkdir(parents=True, exist_ok=True)
        page.screenshot(path=str(output_dir / "debug_login_failed.png"), full_page=True)
        with open(output_dir / "debug_login_failed.html", "w", encoding="utf-8") as f:
            f.write(page.evaluate("document.body.innerHTML"))
            
        sys.exit(1)

    print("✅  Logged in successfully!")


# ──────────────────────────────────────────────
# 3. Search and Autosave
# ──────────────────────────────────────────────
def search_and_save_jobs(page, keywords, location, time_val, max_pages):
    """Search for jobs matching criteria, click 'Save', and return their URLs."""
    import urllib.parse
    print(f"\n============================================================")
    print(f"   🔍 Autosaving Jobs: '{keywords}' in '{location}'")
    print(f"============================================================")

    query = f"?keywords={urllib.parse.quote(keywords)}&location={urllib.parse.quote(location)}&f_TPR=r{time_val}"
    url = f"https://www.linkedin.com/jobs/search/{query}"
    
    print(f"   ➡️  Navigating to search page...")
    try:
        page.goto(url, wait_until="domcontentloaded", timeout=60000)
    except PlaywrightTimeout:
        print("   ⚠️  Timeout waiting for search page to fully load. Continuing anyway...")
    human_delay(3, 5)

    saved_count = 0
    saved_urls = set()
    
    for page_num in range(1, max_pages + 1):
        print(f"\n   📄 Processing Search Page {page_num}...")
        
        # Scroll down by repeatedly scrolling the last loaded job card into view
        for _ in range(5):
            try:
                cards = page.locator(".job-card-container").all()
                if cards:
                    cards[-1].scroll_into_view_if_needed()
                    human_delay(0.5, 1.0)
            except Exception:
                pass
            
        # Get all job cards on the current page
        job_cards = page.locator(".job-card-container").all()
        print(f"   ✔ Found {len(job_cards)} jobs on this page.")
        
        for idx, card in enumerate(job_cards):
            try:
                card_text = card.inner_text().lower()
                
                # Deduplicate: Skip if already Applied, Archived, or Interviewing
                if any(status in card_text for status in ["applied", "beworben", "archived", "archiviert", "interviewing", "im bewerbungsprozess"]):
                    print(f"      [Job {idx+1}] ⏭️ Skipped (Already Applied/Archived).")
                    continue

                # Click the card to open details on the right
                card.scroll_into_view_if_needed()
                card.click()
                human_delay(1.5, 3.0)
                
                # Get the URL of the job
                job_link = card.locator("a").first.get_attribute("href")
                if job_link:
                    clean_url = "https://www.linkedin.com" + job_link.split("?")[0]
                else:
                    clean_url = None
                
                # In the right pane, look for the save button
                save_btn = page.locator(".jobs-save-button").first
                if save_btn.is_visible(timeout=5000):
                    # Check if the job is in German
                    details_pane = page.locator(".job-view-layout, #job-details, .jobs-details").first
                    details_text = details_pane.inner_text() if details_pane.is_visible() else ""
                    
                    # Also check the title just in case
                    title_text = card.inner_text()
                    
                    if is_german(details_text) or is_german(title_text):
                        print(f"      [Job {idx+1}] 🇩🇪 Skipped (German ad).")
                        continue

                    is_pressed = save_btn.get_attribute("aria-pressed")
                    # If it's not saved yet, click save
                    if is_pressed == "false":
                        save_btn.click()
                        saved_count += 1
                        print(f"      [Job {idx+1}] 💾 Saved new job!")
                        human_delay(1.0, 2.0)
                        if clean_url: saved_urls.add(clean_url)
                    else:
                        print(f"      [Job {idx+1}] ✔ Already saved.")
                        if clean_url: saved_urls.add(clean_url)
                else:
                    print(f"      [Job {idx+1}] ⚠️ Could not find save button.")
            except Exception as e:
                print(f"      [Job {idx+1}] ❌ Error processing job: {e}")
                
        # Try to go to next page
        next_button = page.locator(
            "button.artdeco-pagination__button--next, "
            "button[aria-label='Next'], "
            "button[aria-label*='Next'], "
            "button.jobs-search-pagination__button--next"
        ).first
        
        if next_button.is_visible() and next_button.is_enabled():
            print("   ➡️  Moving to next search page...")
            next_button.click()
            human_delay(3, 5)
        else:
            print("   ✔ No more search pages found.")
            break
            
    print(f"✅  Autosave complete! Saved {saved_count} new jobs.")
    return list(saved_urls)


# ──────────────────────────────────────────────
# 4. Scroll & Collect Job Cards
# ──────────────────────────────────────────────
def _debug_saved_jobs_page(page):
    """Dump the page structure to help debug saved jobs selectors."""
    import os
    output_dir = Path(__file__).resolve().parent / "output"
    output_dir.mkdir(parents=True, exist_ok=True)
    
    screenshot_path = output_dir / "debug_saved_jobs.png"
    page.screenshot(path=str(screenshot_path), full_page=True)
    
    html_path = output_dir / "debug_saved_jobs.html"
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(page.evaluate("document.body.innerHTML"))
        
    print(f"    📸 Saved debug HTML and screenshot for Saved Jobs page to {output_dir}")

    # Dump the main structural elements
    structure = page.evaluate("""() => {
        const info = {};
        info.url = window.location.href;
        info.title = document.title;

        // Find all list items on the page
        const allLi = document.querySelectorAll('li');
        info.totalLiElements = allLi.length;

        // Find elements with 'job' in class or id
        const jobEls = document.querySelectorAll('[class*="job"], [id*="job"]');
        info.jobElements = Array.from(jobEls).slice(0, 20).map(el => ({
            tag: el.tagName,
            id: el.id || '',
            className: (el.className && typeof el.className === 'string') ? el.className.substring(0, 120) : '',
            childCount: el.children.length
        }));

        // Find elements with 'saved' in class or id
        const savedEls = document.querySelectorAll('[class*="saved"], [id*="saved"]');
        info.savedElements = Array.from(savedEls).slice(0, 10).map(el => ({
            tag: el.tagName,
            id: el.id || '',
            className: (el.className && typeof el.className === 'string') ? el.className.substring(0, 120) : '',
            childCount: el.children.length
        }));

        // Find all <ul> and <ol> elements
        const lists = document.querySelectorAll('ul, ol');
        info.lists = Array.from(lists).slice(0, 15).map(el => ({
            tag: el.tagName,
            id: el.id || '',
            className: (el.className && typeof el.className === 'string') ? el.className.substring(0, 120) : '',
            childCount: el.children.length
        }));

        // Find all <a> links that contain '/jobs/view/'
        const jobLinks = document.querySelectorAll('a[href*="/jobs/view/"]');
        info.jobLinkCount = jobLinks.length;
        info.jobLinks = Array.from(jobLinks).slice(0, 5).map(el => ({
            href: el.href,
            text: el.innerText.substring(0, 80).trim(),
            parentClass: (el.parentElement?.className && typeof el.parentElement.className === 'string')
                ? el.parentElement.className.substring(0, 100) : ''
        }));

        return info;
    }""")

    print(f"\n   🔍 DEBUG — Saved Jobs Page Structure:")
    print(f"      URL: {structure.get('url')}")
    print(f"      Title: {structure.get('title')}")
    print(f"      Total <li> elements: {structure.get('totalLiElements')}")
    print(f"      Job links (/jobs/view/): {structure.get('jobLinkCount')}")

    if structure.get('jobLinks'):
        print(f"\n      Sample job links:")
        for link in structure['jobLinks']:
            print(f"        🔗 {link['text'][:60]}")
            print(f"           href: {link['href']}")
            print(f"           parent class: {link['parentClass'][:80]}")

    if structure.get('jobElements'):
        print(f"\n      Elements with 'job' in class/id ({len(structure['jobElements'])}):")
        for el in structure['jobElements'][:10]:
            print(f"        <{el['tag']}> id='{el['id']}' class='{el['className'][:80]}' children={el['childCount']}")

    if structure.get('lists'):
        print(f"\n      Lists (<ul>/<ol>) with children ({len(structure['lists'])}):")
        for el in structure['lists']:
            if el['childCount'] > 0:
                print(f"        <{el['tag']}> id='{el['id']}' class='{el['className'][:80]}' children={el['childCount']}")

    print()
    return structure


# Job card selector — broad set of selectors to match various LinkedIn layouts
JOB_CARD_SELECTORS = ", ".join([
    "li.reusable-search__result-container",
    "li.scaffold-layout__list-item",
    "li.jobs-search-results__list-item",
    "div.job-card-container",
    "li.ember-view.jobs-search-results__list-item",
    "div.job-card-list",
    "li.job-card-container__list-item",
])


def scroll_and_collect_jobs(page):
    """Scroll through the saved-jobs page to load all cards, then return them."""
    print("📂  Navigating to Saved Jobs …")
    page.goto(LINKEDIN_SAVED_JOBS_URL, wait_until="domcontentloaded", timeout=60_000)
    print("    ✔ Page loaded. Waiting for content to render …")
    human_delay(5, 8)

    # Wait for the page to settle
    try:
        page.wait_for_load_state("load", timeout=15_000)
    except PlaywrightTimeout:
        pass

    # Debug: dump what's on the page
    structure = _debug_saved_jobs_page(page)

    # If we found job links, we can proceed even if list selectors don't match
    job_link_count = structure.get('jobLinkCount', 0)
    if job_link_count > 0:
        print(f"   ✔ Detected {job_link_count} job link(s) on page.")
    else:
        print("   ⚠️  No job links detected yet. Will try scrolling …")

    all_job_urls = set()

    for page_num in range(1, 20):  # limit to 20 pages max
        prev_count = 0
        no_change_rounds = 0

        for attempt in range(1, MAX_SCROLL_ATTEMPTS + 1):
            # Scroll to bottom
            page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            pause = random.uniform(SCROLL_PAUSE_MIN, SCROLL_PAUSE_MAX)
            time.sleep(pause)

            # Count unique job links
            current_count = page.evaluate("""() => {
                const links = document.querySelectorAll('a[href*="/jobs/view/"]');
                const uniqueIds = new Set();
                for (const link of links) {
                    const match = link.href.match(/\\/jobs\\/view\\/(\\d+)/);
                    if (match) uniqueIds.add(match[1]);
                }
                return uniqueIds.size;
            }""")

            if current_count > prev_count:
                print(f"   ↻ Scroll #{attempt} on Page {page_num}: loaded {current_count} jobs …")
                prev_count = current_count
                no_change_rounds = 0
            else:
                no_change_rounds += 1

            if no_change_rounds >= 3:
                break

        # Extract URLs for current page
        page_urls = page.evaluate("""() => {
            const links = document.querySelectorAll('a[href*="/jobs/view/"]');
            const urls = new Set();
            for (let i = 0; i < links.length; i++) {
                const match = links[i].href.match(/\\/jobs\\/view\\/(\\d+)/);
                if (match) {
                    // Keep only the clean URL, strip any query parameters
                    urls.add(links[i].href.split('?')[0]);
                }
            }
            return Array.from(urls);
        }""")
        
        for u in page_urls:
            all_job_urls.add(u)
            
        print(f"✅  Found {len(page_urls)} jobs on Page {page_num}. Total so far: {len(all_job_urls)}")

        # Try to find and click the "Next" button
        print("   🔍 Looking for Next page button...")
        next_button = None
        
        # Try generic text and label locators
        try:
            btn1 = page.locator("button:has-text('Next')").first
            if btn1.is_visible() and btn1.is_enabled(): next_button = btn1
        except Exception: pass
        
        if not next_button:
            try:
                btn2 = page.locator("[aria-label*='Next']").first
                if btn2.is_visible() and btn2.is_enabled(): next_button = btn2
            except Exception: pass

        if not next_button:
            try:
                btn3 = page.locator("button.artdeco-pagination__button--next").first
                if btn3.is_visible() and btn3.is_enabled(): next_button = btn3
            except Exception: pass
        
        if next_button:
            print(f"   ➡️  Moving to next page ...")
            try:
                next_button.click()
                human_delay(3, 5)
                # Wait for page to settle
                try:
                    page.wait_for_load_state("domcontentloaded", timeout=15000)
                except Exception:
                    pass
            except Exception as e:
                print(f"   ⚠️  Could not click next page: {e}")
                break
        else:
            print("   ✔ No more pages found.")
            break

    print(f"   ✔ Extracted a grand total of {len(all_job_urls)} unique job URLs.")
    return list(all_job_urls)


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


def extract_job_details(page, job_url, index, total):
    """Navigate to a job URL and scrape all available detail fields."""
    job = {
        "title": "N/A",
        "company": "N/A",
        "job_url": job_url,
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
        page.goto(job_url, wait_until="domcontentloaded", timeout=30_000)
        human_delay(2, 4)

        # Wait for detail pane to load
        try:
            page.wait_for_selector(
                ".jobs-unified-top-card, .job-details-jobs-unified-top-card__job-title, "
                ".jobs-details__main-content, .jobs-description",
                timeout=8_000,
            )
        except PlaywrightTimeout:
            pass

        if index == 0:
            import os
            os.makedirs("output", exist_ok=True)
            html = page.evaluate("document.body.innerHTML")
            with open("output/debug_job_0.html", "w", encoding="utf-8") as f:
                f.write(html)
            page.screenshot(path="output/debug_job_0.png", full_page=True)
            print("    📸 Saved debug HTML and screenshot for first job.")

        # Helper to try multiple selectors
        def get_first_valid_text(selectors):
            for sel in selectors:
                try:
                    el = page.query_selector(sel)
                    if el and el.is_visible():
                        t = safe_text(el)
                        if t and t != "N/A":
                            return t, el
                except Exception:
                    continue
            return "N/A", None

        # ---- Title from detail panel ----
        title_text, _ = get_first_valid_text([
            ".job-details-jobs-unified-top-card__job-title",
            ".jobs-unified-top-card__job-title",
            "h1.t-24",
            "h2.jobs-unified-top-card__job-title",
            ".top-card-layout__title",
            "h1.topcard__title",
            "h1"
        ])
        
        # Bulletproof fallback: use the browser tab title (e.g., "Software Engineer at Google | LinkedIn")
        if title_text == "N/A":
            try:
                full_title = page.title()
                if full_title:
                    # Strip " | LinkedIn" from the end
                    clean_title = full_title.split(" | ")[0]
                    # Sometimes it says "Google hiring Software Engineer..."
                    if " hiring " in clean_title:
                        title_text = clean_title.split(" hiring ")[-1]
                    else:
                        title_text = clean_title
            except Exception:
                pass

        job["title"] = title_text

        # ---- Company from detail panel ----
        company_text, company_el = get_first_valid_text([
            ".job-details-jobs-unified-top-card__company-name a",
            ".jobs-unified-top-card__company-name a",
            "a.jobs-unified-top-card__company-name",
            ".topcard__org-name-link",
            "a[href*='/company/']",
            ".job-details-jobs-unified-top-card__company-name"
        ])
        job["company"] = company_text
        if company_el:
            company_href = safe_attr(company_el, "href")
            if company_href != "N/A":
                if company_href.startswith("/"):
                    company_href = "https://www.linkedin.com" + company_href
                job["company_url"] = company_href.split("?")[0]

        # ---- Location from detail panel ----
        location_text, _ = get_first_valid_text([
            ".job-details-jobs-unified-top-card__bullet",
            ".jobs-unified-top-card__bullet",
            ".jobs-unified-top-card__workplace-type",
            ".topcard__flavor--bullet",
            "span.tvm__text--low-emphasis",
            ".job-details-jobs-unified-top-card__primary-description span"
        ])
        job["location"] = location_text

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
        # Click "Show more" if present to expand the full description
        try:
            show_more_btns = page.locator(
                "button.jobs-description__footer-button, "
                "button.show-more-less-html__button, "
                "button:has-text('Show more'), "
                "button:has-text('Mehr anzeigen')"
            ).all()
            for btn in show_more_btns:
                if btn.is_visible():
                    btn.click()
                    human_delay(0.5, 1.0)
                    break
        except Exception:
            pass

        desc_el = page.query_selector(
            "#job-details, "
            ".jobs-description__content, "
            ".jobs-description-content__text, "
            ".jobs-box__html-content, "
            ".description__text, "
            "article"
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

        # ---- Bulletproof Fallback via innerText & page title ----
        if "N/A" in (job["title"], job["company"], job["location"], job["description"], job["date_posted"]):
            try:
                page_text = page.evaluate("document.body.innerText") or ""
                lines = [line.strip() for line in page_text.split('\n') if line.strip()]
                
                # 1. Fallback for Title & Company via Browser Tab Title
                if job["company"] == "N/A" or job["title"] == "N/A":
                    full_title = page.title()
                    if full_title:
                        clean_title = full_title.split(" | ")[0]
                        if " hiring " in clean_title:
                            c, t = clean_title.split(" hiring ", 1)
                            if job["title"] == "N/A": job["title"] = t
                            if job["company"] == "N/A": job["company"] = c
                        elif " at " in clean_title:
                            parts = clean_title.rsplit(" at ", 1)
                            if job["title"] == "N/A": job["title"] = parts[0]
                            if job["company"] == "N/A": job["company"] = parts[1]

                # 2. Fallback for Company & Location via top lines of page text
                # Usually: Line 0 is Title, Line 1 is Company, Line 2 is Location
                if job["company"] == "N/A" and len(lines) > 1:
                    job["company"] = lines[1]
                
                if job["location"] == "N/A" and len(lines) > 2:
                    job["location"] = lines[2]

                # 3. Fallback for date posted
                if job["date_posted"] == "N/A":
                    date_match = re.search(r'(?:Posted\s+)?(?:vor\s+)?(\d+\s+(?:day|week|month|year|hour|tag|woche|monat|jahr|stunde)(?:s|n|en|e)?(?:\s+ago)?)', page_text, re.IGNORECASE)
                    if date_match:
                        job["date_posted"] = date_match.group(1).strip()

                # 4. Fallback for description
                if job["description"] == "N/A":
                    # Look for text between "About the job" (English/German) and common footers
                    jd_match = re.search(
                        r'(?:About the job|Über diesen Job|Job description)\s*(.*?)\s*(?:Show more|Mehr anzeigen|About the company|Über das Unternehmen|Industry|Branche|Employment type|Beschäftigungsart|Follow|Folgen)', 
                        page_text, 
                        re.IGNORECASE | re.DOTALL
                    )
                    
                    if jd_match:
                        jd = jd_match.group(1).strip()
                        if len(jd) > 50:
                            job["description"] = re.sub(r'\n{3,}', '\n\n', jd)
                    else:
                        # Dump the text for debugging!
                        import os
                        from pathlib import Path
                        debug_path = Path(__file__).resolve().parent / "output" / "debug_page_text.txt"
                        with open(debug_path, "w", encoding="utf-8") as f:
                            f.write(page_text)
                            
                        # Alternative: Heuristic JS to find the job description container (highest text density)
                        try:
                            jd = page.evaluate("""() => {
                                let bestEl = null;
                                let maxScore = 0;
                                document.querySelectorAll('div, section, article, main').forEach(el => {
                                    const text = el.innerText || '';
                                    if (text.length > 300 && text.length < 25000) {
                                        const childrenCount = el.querySelectorAll('*').length;
                                        const score = text.length / (childrenCount + 1);
                                        if (score > maxScore) {
                                            maxScore = score;
                                            bestEl = el;
                                        }
                                    }
                                });
                                return bestEl ? bestEl.innerText : '';
                            }""")
                            if jd and len(jd) > 100:
                                job["description"] = jd
                        except Exception:
                            pass
            except Exception:
                pass

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

    email, password, search_keywords_list, search_location, search_time, autosave_max_pages = load_credentials()

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

        # Step 2 — Autosave jobs (if configured)
        job_urls = []
        if search_keywords_list:
            for keyword in search_keywords_list:
                found_urls = search_and_save_jobs(page, keyword, search_location, search_time, autosave_max_pages)
                if found_urls:
                    job_urls.extend(found_urls)
            # Remove duplicates
            job_urls = list(set(job_urls))
            print(f"\n✅ Total unique jobs collected from search phase: {len(job_urls)}")

        # Step 3 — Scroll & collect URLs (only if we didn't get them from search)
        if not job_urls:
            print("\n   [!] No jobs collected from search, falling back to 'Saved Jobs' tab...")
            job_urls = scroll_and_collect_jobs(page)

        if not job_urls:
            print("❌  No saved jobs found.")
            browser.close()
            sys.exit(1)

        # Step 3 — Extract details
        total = len(job_urls)
        print(f"\n📋  Extracting details from {total} saved jobs …\n")
        jobs = []
        for i, url in enumerate(job_urls):
            job_data = extract_job_details(page, url, i, total)
            jobs.append(job_data)
            human_delay(1.0, 2.5)

        browser.close()

    # Step 4 — Export
    print(f"\n💾  Exporting {len(jobs)} jobs to Excel …")
    output_path = export_to_excel(jobs, args.output)
    print(f"✅  Done!  File saved to: {output_path.resolve()}\n")


if __name__ == "__main__":
    main()
