"""
PHASE 3: Facility Harvester Script — ENHANCED
==============================================
Features:
  - Progress bars (tqdm) with % completion
  - Smart empty-page detection with extended wait + retry before declaring empty
  - Post-scrape interactive menu: re-visit ALL empty or specific institutions by number
  - WebDriver only launches AFTER checkpoint prompt
  - Robust pagination: waits for rows to stabilize before moving on
  - Checkpoint save/load
  - Colored console output

Version: 2.0
"""

import time
import json
import re
import sys
import os
import pandas as pd
from datetime import datetime

# ── Optional colored output ─────────────────────────────────────────────────
try:
    from colorama import init, Fore, Style
    init(autoreset=True)
    def green(s):  return Fore.GREEN  + str(s) + Style.RESET_ALL
    def red(s):    return Fore.RED    + str(s) + Style.RESET_ALL
    def yellow(s): return Fore.YELLOW + str(s) + Style.RESET_ALL
    def cyan(s):   return Fore.CYAN   + str(s) + Style.RESET_ALL
    def bold(s):   return Style.BRIGHT + str(s) + Style.RESET_ALL
except ImportError:
    def green(s):  return str(s)
    def red(s):    return str(s)
    def yellow(s): return str(s)
    def cyan(s):   return str(s)
    def bold(s):   return str(s)

# ── Optional tqdm progress bar ───────────────────────────────────────────────
try:
    from tqdm import tqdm
    TQDM_AVAILABLE = True
except ImportError:
    TQDM_AVAILABLE = False
    print(yellow("[WARN] tqdm not installed. Run: pip install tqdm  (progress bars disabled)"))

# ── Selenium ─────────────────────────────────────────────────────────────────
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    StaleElementReferenceException,
    ElementClickInterceptedException,
)

# ==========================================
# CONFIGURATION
# ==========================================
LOGIN_URL        = "https://www.ocip.express/"
TARGET_URL       = "https://www.ocip.express/FacilityAdmin/Index"
OUTPUT_JSON      = "facilities_master_list.json"
OUTPUT_EXCEL     = "facilities_master_list.xlsx"
CHECKPOINT_FILE  = "phase3_checkpoint.json"

# ── Timing (seconds) ──────────────────────────────────────────────────────────
PAGE_LOAD_WAIT          = 1.5   # after selecting institution, base wait
PAGINATION_WAIT         = 2.0   # after clicking next page
DROPDOWN_CLOSE_WAIT     = 0.7
LOADING_MASK_TIMEOUT    = 15    # how long to wait for k-loading-mask to vanish

# ── Empty-page detection ──────────────────────────────────────────────────────
# If 0 rows found immediately, we wait EMPTY_RETRY_WAIT seconds and try again,
# up to EMPTY_MAX_RETRIES times before giving up and declaring the page empty.
EMPTY_RETRY_WAIT   = 0.5   # seconds to wait between retries
EMPTY_MAX_RETRIES  = 3     # total extra attempts

# ── Row stabilisation ─────────────────────────────────────────────────────────
# After navigating to a page we poll until the row count stops changing.
ROW_STABLE_POLLS      = 2      # how many consecutive same-count reads before "stable"
ROW_STABLE_INTERVAL   = 1.0    # seconds between polls
ROW_STABLE_TIMEOUT    = 3     # give up after this many seconds total

# ==========================================
# DRIVER SETUP
# ==========================================
def get_driver():
    """Initialize Chrome WebDriver with optimal settings."""
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    options.add_experimental_option("detach", True)
    driver = webdriver.Chrome(options=options)
    driver.execute_script(
        "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    )
    return driver

# ==========================================
# UTILITY: Loading / Waiting
# ==========================================
def wait_for_loading_complete(driver, timeout=LOADING_MASK_TIMEOUT):
    """Wait for Kendo loading masks to vanish."""
    try:
        time.sleep(0.4)
        WebDriverWait(driver, timeout).until(
            EC.invisibility_of_element_located((By.CSS_SELECTOR, ".k-loading-mask"))
        )
        time.sleep(0.3)
    except Exception:
        pass


def wait_for_rows_stable(driver, timeout=ROW_STABLE_TIMEOUT):
    """
    Poll the row count until it stops changing for ROW_STABLE_POLLS consecutive
    reads, or timeout expires.  Returns the stable list of rows.
    """
    deadline  = time.time() + timeout
    last_count = -1
    streak     = 0

    while time.time() < deadline:
        try:
            rows = driver.find_elements(By.CSS_SELECTOR, "tr.k-master-row")
        except Exception:
            rows = []
        count = len(rows)

        if count == last_count and count > 0:
            streak += 1
            if streak >= ROW_STABLE_POLLS:
                return rows          # stable non-empty result
        else:
            streak = 0
            last_count = count

        time.sleep(ROW_STABLE_INTERVAL)

    # Return whatever we have
    try:
        return driver.find_elements(By.CSS_SELECTOR, "tr.k-master-row")
    except Exception:
        return []


def wait_for_rows_with_retry(driver, institution_name):
    """
    Try to get rows.  If none found, wait EMPTY_RETRY_WAIT and retry up to
    EMPTY_MAX_RETRIES times before concluding the institution is truly empty.
    Returns (rows, declared_empty: bool)
    """
    # First attempt: wait for stability
    rows = wait_for_rows_stable(driver)
    if rows:
        return rows, False

    # Extended retry loop for slow-loading pages
    for attempt in range(1, EMPTY_MAX_RETRIES + 1):
        print(yellow(f"         ⏳ No rows yet — waiting {EMPTY_RETRY_WAIT}s "
                     f"(attempt {attempt}/{EMPTY_MAX_RETRIES}) for {institution_name}…"))
        time.sleep(EMPTY_RETRY_WAIT)
        wait_for_loading_complete(driver)

        rows = wait_for_rows_stable(driver)
        if rows:
            print(green(f"         ✓ Rows appeared after {attempt} extra wait(s)"))
            return rows, False

        # Also check pagination info — data might be present but rows styled differently
        start, end, total = parse_pagination_info(driver)
        if total > 0:
            # Data exists, try once more with a long wait
            print(yellow(f"         ℹ Pagination says {total} items — waiting extra…"))
            time.sleep(EMPTY_RETRY_WAIT * 2)
            wait_for_loading_complete(driver)
            rows = wait_for_rows_stable(driver)
            if rows:
                return rows, False

    return [], True  # genuinely empty after all retries

# ==========================================
# UTILITY: Pagination
# ==========================================
def parse_pagination_info(driver):
    """Return (start, end, total) from pager label, e.g. '1 - 50 of 120 items'."""
    try:
        pager = driver.find_element(By.CSS_SELECTOR, "span.k-pager-info.k-label")
        text  = pager.text.strip()
        m = re.match(r'(\d+)\s*-\s*(\d+)\s+of\s+(\d+)\s+items?', text)
        if m:
            return int(m.group(1)), int(m.group(2)), int(m.group(3))
    except Exception:
        pass
    return 0, 0, 0


def has_next_page(driver):
    try:
        btn = driver.find_element(
            By.CSS_SELECTOR, "a.k-pager-nav[aria-label='Go to the next page']"
        )
        return btn.get_attribute("aria-disabled") != "true"
    except Exception:
        return False


def click_next_page(driver, wait):
    try:
        btn = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "a.k-pager-nav[aria-label='Go to the next page']")
        ))
        if btn.get_attribute("aria-disabled") == "true":
            return False
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
        time.sleep(0.3)
        try:
            btn.click()
        except Exception:
            driver.execute_script("arguments[0].click();", btn)
        wait_for_loading_complete(driver)
        time.sleep(PAGINATION_WAIT)
        return True
    except Exception:
        return False


def reset_to_first_page(driver, wait):
    try:
        start, _, _ = parse_pagination_info(driver)
        if start <= 1:
            return True
        btn = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR,
             "a.k-pager-nav.k-pager-first[aria-label='Go to the first page']")
        ))
        if btn.get_attribute("aria-disabled") == "true":
            return True
        driver.execute_script("arguments[0].click();", btn)
        wait_for_loading_complete(driver)
        time.sleep(PAGINATION_WAIT)
        return True
    except Exception:
        return False

# ==========================================
# CHECKPOINT
# ==========================================
def save_checkpoint(data, current_index, institution_names, empty_institutions):
    checkpoint = {
        "timestamp":            datetime.now().isoformat(),
        "current_index":        current_index,
        "total_institutions":   len(institution_names),
        "facilities_collected": len(data),
        "institution_names":    institution_names,
        "empty_institutions":   empty_institutions,   # list of {index, name}
        "data":                 data,
    }
    with open(CHECKPOINT_FILE, 'w', encoding='utf-8') as f:
        json.dump(checkpoint, f, indent=2)


def load_checkpoint():
    try:
        with open(CHECKPOINT_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return None

# ==========================================
# MODULE 1: Get Institution Names
# ==========================================
def get_institution_names(driver, wait):
    print("\n" + "=" * 55)
    print(bold("  GATHERING INSTITUTION LIST"))
    print("=" * 55)
    try:
        trigger = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "span[aria-controls$='listbox']")
        ))
        trigger.click()

        listbox = wait.until(EC.visibility_of_element_located(
            (By.CSS_SELECTOR, "ul[id$='listbox'][aria-hidden='false']")
        ))
        time.sleep(0.6)

        options = listbox.find_elements(By.TAG_NAME, "li")
        names = []
        for opt in options:
            text = opt.text.strip()
            if text and "Select HEI" not in text:
                names.append(text)

        print(green(f"  ✓ Found {len(names)} institutions"))
        driver.find_element(By.TAG_NAME, "body").click()
        time.sleep(DROPDOWN_CLOSE_WAIT)
        return names

    except Exception as e:
        print(red(f"  ✗ FAILED: {e}"))
        return []

# ==========================================
# MODULE 2: Select Institution
# ==========================================
def select_institution(driver, wait, target_name):
    try:
        trigger = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "span[aria-controls$='listbox']")
        ))
        trigger.click()

        listbox = wait.until(EC.visibility_of_element_located(
            (By.CSS_SELECTOR, "ul[id$='listbox'][aria-hidden='false']")
        ))
        time.sleep(0.4)

        for opt in listbox.find_elements(By.TAG_NAME, "li"):
            if target_name in opt.text:
                driver.execute_script("arguments[0].click();", opt)
                wait_for_loading_complete(driver)
                time.sleep(PAGE_LOAD_WAIT)
                return True

        driver.find_element(By.TAG_NAME, "body").click()
        return False

    except Exception:
        try:
            driver.find_element(By.TAG_NAME, "body").click()
        except Exception:
            pass
        return False

# ==========================================
# MODULE 3: Scrape One Page of Rows
# ==========================================
def scrape_rows(rows, institution_name):
    """
    Extract data from a stable list of tr.k-master-row elements.
    Looks at the screenshot: columns are
      0: Branch/HEI abbrev  1: Facility Name  2: Facility Type
      3: Experts?  4: Equipment?  5: Enabled  6: Actions (Manage link)
    Adjust indices if the live page differs.
    """
    page_data = []

    for row in rows:
        try:
            cells = row.find_elements(By.TAG_NAME, "td")
            if not cells:
                continue

            def cell_text(idx):
                try:
                    return cells[idx].text.strip()
                except Exception:
                    return ""

            branch        = cell_text(0)
            facility_name = cell_text(1)
            facility_type = cell_text(2)

            # --- Manage / Actions link ---
            manage_url = "Not Found"

            # Strategy 1: last cell contains the link
            try:
                link = cells[-1].find_element(By.TAG_NAME, "a")
                manage_url = link.get_attribute("href") or "Not Found"
            except Exception:
                pass

            # Strategy 2: scan all cells for href containing 'Manage' or 'Details'
            if manage_url == "Not Found":
                try:
                    for cell in cells:
                        links = cell.find_elements(By.TAG_NAME, "a")
                        for lnk in links:
                            href = lnk.get_attribute("href") or ""
                            title = lnk.get_attribute("title") or lnk.text or ""
                            if any(kw in href or kw in title
                                   for kw in ["Manage", "Details", "Edit", "View"]):
                                manage_url = href
                                break
                        if manage_url != "Not Found":
                            break
                except Exception:
                    pass

            page_data.append({
                "Institution":   institution_name,
                "Branch":        branch,
                "Facility_Name": facility_name,
                "Facility_Type": facility_type,
                "Manage_URL":    manage_url,
                "Scraped_At":    datetime.now().isoformat(),
            })

        except StaleElementReferenceException:
            continue
        except Exception:
            continue

    return page_data

# ==========================================
# MODULE 4: Full Pagination Loop
# ==========================================
def scrape_all_pages_for_institution(driver, wait, institution_name):
    """
    Scrapes all pages for one institution.
    Uses wait_for_rows_with_retry so slow-loading pages are not skipped.
    """
    all_facilities = []
    current_page   = 1

    # ── First page: use the extended retry logic ──────────────────────────────
    rows, is_empty = wait_for_rows_with_retry(driver, institution_name)

    if is_empty:
        # One final check via pagination label
        _, _, total = parse_pagination_info(driver)
        if total == 0:
            return []          # Genuinely no data
        # Pagination says there IS data — wait a bit more
        print(yellow(f"         ⚠ Pagination says {total} items but no rows visible. "
                     f"Waiting extra 10s…"))
        time.sleep(10)
        wait_for_loading_complete(driver)
        rows = wait_for_rows_stable(driver)
        if not rows:
            return []

    _, _, total = parse_pagination_info(driver)
    print(cyan(f"      → Detected {total if total else '?'} total facilities"))

    # Progress bar for pages (estimate)
    pages_estimate = max(1, (total // 50) + 1) if total else 999

    page_bar = None
    if TQDM_AVAILABLE and total:
        page_bar = tqdm(
            total=total,
            desc=f"      {institution_name[:30]}",
            unit="row",
            leave=False,
            bar_format="{l_bar}{bar}| {n_fmt}/{total_fmt} rows [{elapsed}<{remaining}]"
        )

    while True:
        page_data = scrape_rows(rows, institution_name)
        all_facilities.extend(page_data)

        if page_bar:
            page_bar.update(len(page_data))

        print(f"         Page {current_page}: {green(str(len(page_data)))} records  "
              f"(running total: {len(all_facilities)})")

        # Check if we should move to next page
        start, end, total_now = parse_pagination_info(driver)
        if not has_next_page(driver) or (total_now > 0 and end >= total_now):
            break

        if not click_next_page(driver, wait):
            break

        # Wait for the *new* page rows to stabilise before scraping
        rows = wait_for_rows_stable(driver)
        if not rows:
            # Try once more with extra time
            time.sleep(EMPTY_RETRY_WAIT)
            wait_for_loading_complete(driver)
            rows = wait_for_rows_stable(driver)
            if not rows:
                print(yellow("         ⚠ No rows on this page after waiting — stopping pagination"))
                break

        current_page += 1
        if current_page > 200:   # absolute safety
            break

    if page_bar:
        page_bar.close()

    if current_page > 1:
        reset_to_first_page(driver, wait)

    return all_facilities

# ==========================================
# SCRAPE ONE INSTITUTION (helper)
# ==========================================
def scrape_institution(driver, wait, uni_name, idx, total_count):
    """Select an institution and scrape all its pages. Returns list of records."""
    print(f"\n[{idx + 1}/{total_count}] {bold(uni_name)}")

    if not select_institution(driver, wait, uni_name):
        print(red("      ✗ Selection failed — skipping"))
        return None   # None = selection error, distinct from []

    data = scrape_all_pages_for_institution(driver, wait, uni_name)
    return data

# ==========================================
# POST-SCRAPE INTERACTIVE MENU
# ==========================================
def post_scrape_menu(driver, wait, master_list, empty_institutions, institution_names):
    """
    After the main scrape, ask the user if they want to re-visit empties.
    Allows: re-visit ALL empties, specific ones by number, or finish.
    """
    if not empty_institutions:
        print(green("\n✓ No empty institutions detected — all done!"))
        return master_list

    print("\n" + "=" * 60)
    print(bold("  POST-SCRAPE: EMPTY INSTITUTIONS REVIEW"))
    print("=" * 60)
    print(yellow(f"  {len(empty_institutions)} institution(s) had NO data collected:\n"))

    for item in empty_institutions:
        print(f"    [{item['index'] + 1:>3}] {item['name']}")

    while True:
        print("\n" + "-" * 60)
        print("  Options:")
        print("    A  — Re-visit ALL empty institutions")
        print("    S  — Re-visit SPECIFIC institutions (enter numbers)")
        print("    Q  — Quit / finish without re-visiting")
        choice = input("\n  Your choice: ").strip().upper()

        if choice == "Q":
            print(cyan("  → Finishing without re-visit."))
            break

        elif choice == "A":
            targets = empty_institutions[:]
            newly_empty = _revisit_targets(driver, wait, master_list, targets,
                                           institution_names, empty_institutions)
            empty_institutions = newly_empty
            if not empty_institutions:
                print(green("  ✓ All previously-empty institutions now have data!"))
                break

        elif choice == "S":
            raw = input(
                "  Enter institution numbers from the list above "
                "(comma-separated, e.g. 3,7,12): "
            ).strip()
            try:
                chosen_display_nums = [int(x.strip()) for x in raw.split(",") if x.strip()]
            except ValueError:
                print(red("  ✗ Invalid input — please enter numbers only"))
                continue

            # Map display numbers back to empty_institutions entries
            targets = []
            for dn in chosen_display_nums:
                match = [e for e in empty_institutions if e["index"] + 1 == dn]
                if match:
                    targets.append(match[0])
                else:
                    print(yellow(f"  ⚠ Number {dn} not in empty list — skipping"))

            if not targets:
                print(red("  ✗ No valid selections — try again"))
                continue

            newly_empty = _revisit_targets(driver, wait, master_list, targets,
                                           institution_names, empty_institutions)
            # Remove re-scraped ones from empty list
            revisited_names = {t["name"] for t in targets}
            still_empty = [e for e in empty_institutions if e["name"] in newly_empty_names(newly_empty)]
            empty_institutions = still_empty
            if not empty_institutions:
                print(green("  ✓ No more empty institutions!"))
                break
        else:
            print(red("  ✗ Unrecognised option — please enter A, S, or Q"))

    return master_list


def newly_empty_names(newly_empty):
    return {e["name"] for e in newly_empty}


def _revisit_targets(driver, wait, master_list, targets, institution_names, old_empty):
    """Re-scrape target institutions. Returns list of still-empty entries."""
    print(cyan(f"\n  → Re-visiting {len(targets)} institution(s)…"))
    still_empty = []

    bar = None
    if TQDM_AVAILABLE:
        bar = tqdm(targets, desc="  Re-scraping", unit="inst")

    iterable = bar if bar else targets
    for item in iterable:
        uni_name = item["name"]
        idx      = item["index"]

        # Remove any previously collected (possibly partial) records for this institution
        before = len(master_list)
        master_list[:] = [r for r in master_list if r.get("Institution") != uni_name]
        removed = before - len(master_list)
        if removed:
            print(yellow(f"      ℹ Removed {removed} old records for {uni_name} before re-scrape"))

        data = scrape_institution(driver, wait, uni_name, idx, len(institution_names))

        if data is None:
            print(red(f"      ✗ Selection error for {uni_name}"))
            still_empty.append(item)
        elif len(data) == 0:
            print(yellow(f"      ⚠ Still no data for {uni_name}"))
            still_empty.append(item)
        else:
            master_list.extend(data)
            print(green(f"      ✓ Collected {len(data)} records for {uni_name}"))

        save_checkpoint(master_list, idx + 1, institution_names,
                        [e for e in old_empty if e["name"] != uni_name])
        time.sleep(0.5)

    if bar:
        bar.close()

    return still_empty

# ==========================================
# SAVE RESULTS
# ==========================================
def save_results(master_list):
    print("\n" + "=" * 60)
    print(bold("  SAVING RESULTS"))
    print("=" * 60)

    with open(OUTPUT_JSON, 'w', encoding='utf-8') as f:
        json.dump(master_list, f, indent=2, ensure_ascii=False)
    print(green(f"  ✓ JSON  → {OUTPUT_JSON}  ({len(master_list)} records)"))

    try:
        df = pd.DataFrame(master_list)
        df.to_excel(OUTPUT_EXCEL, index=False)
        print(green(f"  ✓ Excel → {OUTPUT_EXCEL}"))
    except Exception as e:
        print(yellow(f"  ⚠ Excel save failed: {e}"))

# ==========================================
# MAIN
# ==========================================
def main():
    print("\n" + "=" * 60)
    print(bold("  PHASE 3: FACILITY METADATA HARVESTER  v2.0"))
    print("=" * 60)

    # ── Checkpoint prompt BEFORE opening browser ──────────────────────────────
    master_list        = []
    start_index        = 0
    institution_names  = []
    empty_institutions = []

    checkpoint = load_checkpoint()
    resume = False

    if checkpoint:
        print(yellow(f"\n  ⚠  Checkpoint found:"))
        print(f"      Timestamp  : {checkpoint.get('timestamp','?')}")
        print(f"      Progress   : {checkpoint.get('current_index','?')} / "
              f"{checkpoint.get('total_institutions','?')} institutions")
        print(f"      Records    : {checkpoint.get('facilities_collected','?')}")
        ans = input("\n  Resume from checkpoint? (y/n): ").strip().lower()
        if ans == 'y':
            master_list        = checkpoint.get('data', [])
            start_index        = checkpoint.get('current_index', 0)
            institution_names  = checkpoint.get('institution_names', [])
            empty_institutions = checkpoint.get('empty_institutions', [])
            resume = True
            print(green(f"  ✓ Resuming from institution #{start_index + 1}"))
        else:
            print(cyan("  → Starting fresh"))
    else:
        print(cyan("  No checkpoint found — starting fresh"))

    # ── NOW open the browser ──────────────────────────────────────────────────
    print(cyan("\n  Opening browser…"))
    driver = get_driver()
    wait   = WebDriverWait(driver, 20)

    try:
        # Step 1: Login
        driver.get(LOGIN_URL)
        print(bold("\nSTEP 1: Please log in to the portal."))
        input("  >>> Press ENTER when fully logged in… ")

        # Step 2: Navigate to Facilities
        driver.get(TARGET_URL)
        print(cyan(f"\n  Navigating to: {TARGET_URL}"))
        wait_for_loading_complete(driver, timeout=20)
        time.sleep(PAGE_LOAD_WAIT)

        # Step 3: Get institutions (if not loaded from checkpoint)
        if not institution_names:
            institution_names = get_institution_names(driver, wait)
            if not institution_names:
                print(red("  ✗ No institutions found — aborting"))
                return

        total_count = len(institution_names)

        # Step 4: Main scrape loop
        print(bold(f"\nSTEP 2: SCRAPING FACILITIES  ({total_count} institutions)\n"))

        outer_bar = None
        if TQDM_AVAILABLE:
            outer_bar = tqdm(
                total=total_count,
                initial=start_index,
                desc="  Institutions",
                unit="inst",
                bar_format="{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}]"
            )

        for i, uni_name in enumerate(institution_names[start_index:], start=start_index):

            data = scrape_institution(driver, wait, uni_name, i, total_count)

            if data is None:
                # Selection failed — treat as needing re-visit
                empty_institutions.append({"index": i, "name": uni_name, "reason": "selection_failed"})
            elif len(data) == 0:
                print(yellow(f"      ⚠ No data collected for {uni_name}"))
                empty_institutions.append({"index": i, "name": uni_name, "reason": "no_data"})
            else:
                master_list.extend(data)
                print(green(f"      ✓ {len(data)} records  |  Running total: {len(master_list)}"))

            if outer_bar:
                outer_bar.update(1)

            save_checkpoint(master_list, i + 1, institution_names, empty_institutions)
            time.sleep(0.5)

        if outer_bar:
            outer_bar.close()

        # Step 5: Summary
        print("\n" + "=" * 60)
        print(bold("  SCRAPE COMPLETE — SUMMARY"))
        print("=" * 60)
        print(f"  Total institutions  : {total_count}")
        print(green(f"  Records collected   : {len(master_list)}"))
        if empty_institutions:
            print(yellow(f"  Empty institutions  : {len(empty_institutions)}"))
        else:
            print(green("  Empty institutions  : 0"))

        # Step 6: Save
        save_results(master_list)

        # Step 7: Post-scrape interactive re-visit menu
        master_list = post_scrape_menu(
            driver, wait, master_list, empty_institutions, institution_names
        )

        # Save final results again in case re-visits added data
        save_results(master_list)
        print(green("\n  ✓ All done!"))

    except KeyboardInterrupt:
        print(yellow("\n  ⚠  Interrupted by user — saving progress…"))
        save_checkpoint(master_list, start_index, institution_names, empty_institutions)
        save_results(master_list)

    except Exception as e:
        print(red(f"\n  ✗ CRITICAL ERROR: {e}"))
        import traceback
        traceback.print_exc()
        save_checkpoint(master_list, start_index, institution_names, empty_institutions)
        save_results(master_list)

    finally:
        print(cyan("\n  Script finished."))


if __name__ == "__main__":
    main()
