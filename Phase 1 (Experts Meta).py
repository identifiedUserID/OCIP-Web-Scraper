"""
PHASE 1: Expert Harvester Script
================================
This script scrapes expert metadata from the eCampusOntario OCIP Express portal.
It handles dropdown selection, pagination, and saves results to JSON and Excel.

Author: AI Assistant
Version: 2.1 (with pagination reset fix)
"""

import time
import json
import re
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    StaleElementReferenceException,
    ElementClickInterceptedException
)
from datetime import datetime

# ==========================================
# CONFIGURATION
# ==========================================
LOGIN_URL = "https://www.ocip.express/"
OUTPUT_JSON = "experts_master_list.json"
OUTPUT_EXCEL = "experts_master_list.xlsx"
CHECKPOINT_FILE = "checkpoint_progress.json"

# Timing Configuration (adjust if site is slow)
PAGE_LOAD_WAIT = 2.0  # Wait after selecting institution
PAGINATION_WAIT = 1.5  # Wait after clicking next page
DROPDOWN_CLOSE_WAIT = 0.5  # Wait after closing dropdown
LOADING_MASK_TIMEOUT = 10  # Max wait for loading spinner


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
    options.add_experimental_option("detach", True)  # Keep browser open after script ends

    driver = webdriver.Chrome(options=options)

    # Make selenium less detectable
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

    return driver


# ==========================================
# UTILITY FUNCTIONS
# ==========================================
def wait_for_loading_complete(driver, timeout=LOADING_MASK_TIMEOUT):
    """Wait for any loading masks/spinners to disappear."""
    try:
        # First, give the mask a moment to appear
        time.sleep(0.3)

        # Wait for loading mask to disappear
        WebDriverWait(driver, timeout).until(
            EC.invisibility_of_element_located((By.CSS_SELECTOR, ".k-loading-mask"))
        )
    except TimeoutException:
        print("      [Warning] Loading mask timeout - proceeding anyway")
    except:
        pass  # No loading mask appeared


def parse_pagination_info(driver):
    """
    Parse the pagination info text to get current range and total count.
    Returns: (current_start, current_end, total_count) or (0, 0, 0) if not found

    Example: "1 - 100 of 163 items" -> (1, 100, 163)
    """
    try:
        pager_info = driver.find_element(By.CSS_SELECTOR, "span.k-pager-info.k-label")
        info_text = pager_info.text.strip()

        # Parse "1 - 100 of 163 items" format
        match = re.match(r'(\d+)\s*-\s*(\d+)\s+of\s+(\d+)\s+items?', info_text)

        if match:
            start = int(match.group(1))
            end = int(match.group(2))
            total = int(match.group(3))
            return (start, end, total)
        else:
            # Try alternate format "No items to display"
            if "no items" in info_text.lower() or "0 items" in info_text.lower():
                return (0, 0, 0)
            print(f"      [Warning] Could not parse pagination: '{info_text}'")
            return (0, 0, 0)

    except NoSuchElementException:
        print("      [Warning] Pagination info element not found")
        return (0, 0, 0)
    except Exception as e:
        print(f"      [Warning] Pagination parse error: {e}")
        return (0, 0, 0)


def has_next_page(driver):
    """Check if there's a next page available."""
    try:
        next_btn = driver.find_element(By.CSS_SELECTOR, "a.k-pager-nav[aria-label='Go to the next page']")
        is_disabled = next_btn.get_attribute("aria-disabled")
        return is_disabled != "true"
    except NoSuchElementException:
        return False
    except Exception:
        return False


def click_next_page(driver, wait):
    """Click the next page button and wait for table to reload."""
    try:
        next_btn = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "a.k-pager-nav[aria-label='Go to the next page']")
        ))

        # Check if disabled
        if next_btn.get_attribute("aria-disabled") == "true":
            return False

        # Scroll into view and click
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_btn)
        time.sleep(0.2)

        # Try regular click first, fallback to JS click
        try:
            next_btn.click()
        except ElementClickInterceptedException:
            driver.execute_script("arguments[0].click();", next_btn)

        # Wait for new data to load
        wait_for_loading_complete(driver)
        time.sleep(PAGINATION_WAIT)

        return True

    except TimeoutException:
        print("      [Warning] Next page button not clickable")
        return False
    except Exception as e:
        print(f"      [Warning] Next page click failed: {e}")
        return False


def reset_to_first_page(driver, wait):
    """
    Navigate back to the first page of results.
    Returns True if successful or already on first page, False otherwise.
    """
    try:
        # Check if we're already on page 1
        start, end, total = parse_pagination_info(driver)
        if start <= 1:
            return True  # Already on first page

        # Find the "Go to first page" button
        first_page_btn = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "a.k-pager-nav.k-pager-first[aria-label='Go to the first page']")
        ))

        # Check if disabled
        if first_page_btn.get_attribute("aria-disabled") == "true":
            return True  # Already on first page

        # Scroll into view and click
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", first_page_btn)
        time.sleep(0.2)

        # Try regular click first, fallback to JS click
        try:
            first_page_btn.click()
        except ElementClickInterceptedException:
            driver.execute_script("arguments[0].click();", first_page_btn)

        # Wait for new data to load
        wait_for_loading_complete(driver)
        time.sleep(PAGINATION_WAIT)

        print("         → Reset to page 1")
        return True

    except TimeoutException:
        print("         [Warning] First page button not found/clickable")
        return False
    except Exception as e:
        print(f"         [Warning] Failed to reset to first page: {e}")
        return False


def save_checkpoint(data, current_index, institution_names):
    """Save progress checkpoint in case of crash."""
    checkpoint = {
        "timestamp": datetime.now().isoformat(),
        "current_index": current_index,
        "total_institutions": len(institution_names),
        "experts_collected": len(data),
        "data": data
    }
    with open(CHECKPOINT_FILE, 'w', encoding='utf-8') as f:
        json.dump(checkpoint, f, indent=2)


def load_checkpoint():
    """Load previous checkpoint if exists."""
    try:
        with open(CHECKPOINT_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        return None
    except json.JSONDecodeError:
        return None


# ==========================================
# MODULE 1: Get Institution Names
# ==========================================
def get_institution_names(driver, wait):
    """
    Extract all institution names from the dropdown.
    Returns a list of institution name strings.
    """
    print("\n" + "=" * 50)
    print("GATHERING INSTITUTION LIST")
    print("=" * 50)

    try:
        # Open dropdown by clicking the arrow/container
        dropdown_trigger = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "span[aria-controls='HeiId_listbox']")
        ))
        dropdown_trigger.click()

        # Wait for listbox to appear
        listbox = wait.until(EC.visibility_of_element_located((By.ID, "HeiId_listbox")))
        time.sleep(0.5)  # Allow list to fully populate

        # Get all list items
        options = listbox.find_elements(By.TAG_NAME, "li")

        names = []
        for opt in options:
            text = opt.text.strip()
            # Filter out placeholder text
            if text and "Select HEI" not in text and text != "":
                names.append(text)

        print(f"✓ Found {len(names)} institutions to process")

        # Close dropdown
        driver.find_element(By.TAG_NAME, "body").click()
        time.sleep(DROPDOWN_CLOSE_WAIT)

        return names

    except TimeoutException:
        print("✗ FAILED: Dropdown not found or not clickable")
        return []
    except Exception as e:
        print(f"✗ FAILED: {e}")
        return []


# ==========================================
# MODULE 2: Select Institution
# ==========================================
def select_institution(driver, wait, target_name):
    """
    Select an institution from the dropdown using robust matching.
    Returns True if successful, False otherwise.
    """
    try:
        # 1. Open Dropdown
        dropdown_trigger = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "span[aria-controls='HeiId_listbox']")
        ))
        dropdown_trigger.click()

        # 2. Wait for listbox
        listbox = wait.until(EC.visibility_of_element_located((By.ID, "HeiId_listbox")))
        time.sleep(0.3)

        # 3. Find matching option using Python loop (handles hidden chars)
        all_options = listbox.find_elements(By.TAG_NAME, "li")

        target_element = None
        for opt in all_options:
            opt_text = opt.text.strip()
            # Use 'in' comparison to handle minor text differences
            if target_name in opt_text or opt_text in target_name:
                target_element = opt
                break

        if not target_element:
            print(f"      ✗ Could not find '{target_name}' in dropdown")
            driver.find_element(By.TAG_NAME, "body").click()
            return False

        # 4. Click using JavaScript (more reliable)
        driver.execute_script("arguments[0].click();", target_element)

        # 5. Wait for table to load
        wait_for_loading_complete(driver)
        time.sleep(PAGE_LOAD_WAIT)

        return True

    except TimeoutException:
        print(f"      ✗ Timeout selecting '{target_name}'")
        try:
            driver.find_element(By.TAG_NAME, "body").click()
        except:
            pass
        return False
    except Exception as e:
        print(f"      ✗ Selection error: {e}")
        try:
            driver.find_element(By.TAG_NAME, "body").click()
        except:
            pass
        return False


# ==========================================
# MODULE 3: Scrape Single Page of Table
# ==========================================
def scrape_current_page(driver, institution_name):
    """
    Scrape all expert rows from the currently displayed table page.
    Returns a list of expert dictionaries.
    """
    try:
        rows = driver.find_elements(By.CSS_SELECTOR, "tr.k-master-row")
    except:
        return []

    if not rows:
        return []

    page_data = []

    for row in rows:
        try:
            cells = row.find_elements(By.TAG_NAME, "td")

            if len(cells) < 9:
                continue

            # Extract visible data from correct cell indices
            # Based on HTML: cells[6]=Type, cells[7]=Position, cells[8]=Name
            expert_type = cells[6].text.strip()
            position = cells[7].text.strip()
            name = cells[8].text.strip()

            # Extract hidden cell data (for additional context)
            try:
                expert_id = cells[1].text.strip() if len(cells) > 1 else ""
                facility_name = cells[4].text.strip() if len(cells) > 4 else ""
            except:
                expert_id = ""
                facility_name = ""

            # Extract the "Manage" link URL
            manage_url = ""
            profile_url = ""
            try:
                manage_link = row.find_element(By.CSS_SELECTOR, "a[title='View Full Details']")
                manage_url = manage_link.get_attribute("href")
            except NoSuchElementException:
                manage_url = "Not Found"

            try:
                profile_link = row.find_element(By.CSS_SELECTOR, "a[title='View Profile']")
                profile_url = profile_link.get_attribute("href")
            except NoSuchElementException:
                profile_url = "Not Found"

            # Build expert record
            expert_record = {
                "Institution": institution_name,
                "Name": name,
                "Expert_Type": expert_type,
                "Position": position,
                "Facility": facility_name,
                "Expert_ID": expert_id,
                "Manage_URL": manage_url,
                "Profile_URL": profile_url,
                "Scraped_At": datetime.now().isoformat()
            }

            page_data.append(expert_record)

        except StaleElementReferenceException:
            continue
        except Exception as e:
            continue

    return page_data


# ==========================================
# MODULE 4: Scrape All Pages for Institution
# ==========================================
def scrape_all_pages_for_institution(driver, wait, institution_name):
    """
    Scrape ALL pages of experts for a given institution.
    Handles pagination automatically and resets to first page if multiple pages were scraped.
    Returns complete list of experts.
    """
    all_experts = []
    current_page = 1

    # Get initial pagination info
    start, end, total = parse_pagination_info(driver)

    if total == 0:
        print(f"      → No experts found")
        return []

    total_pages = (total + 99) // 100  # Calculate expected pages (ceiling division)
    print(f"      → Found {total} experts across ~{total_pages} page(s)")

    while True:
        # Scrape current page
        page_experts = scrape_current_page(driver, institution_name)
        experts_on_page = len(page_experts)
        all_experts.extend(page_experts)

        print(f"         Page {current_page}: Scraped {experts_on_page} experts (Total so far: {len(all_experts)})")

        # Check if more pages exist
        start, end, total = parse_pagination_info(driver)

        if end >= total:
            # We've reached the last page
            break

        if not has_next_page(driver):
            # No next button available
            break

        # Click next page
        if click_next_page(driver, wait):
            current_page += 1
        else:
            print(f"         [Warning] Could not navigate to page {current_page + 1}")
            break

        # Safety limit to prevent infinite loops
        if current_page > 50:
            print("         [Warning] Safety limit reached (50 pages)")
            break

    # **FIX: Reset to first page if we scraped multiple pages**
    if current_page > 1:
        print(f"      → Institution had {current_page} pages, resetting pagination...")
        reset_to_first_page(driver, wait)

    return all_experts


# ==========================================
# MAIN EXECUTION
# ==========================================
def main():
    """Main execution function."""

    print("\n" + "=" * 60)
    print("  PHASE 1: EXPERT HARVESTER")
    print("  eCampusOntario OCIP Express Portal Scraper")
    print("=" * 60)

    # Initialize
    driver = get_driver()
    wait = WebDriverWait(driver, 20)
    master_list = []
    start_index = 0

    # Check for existing checkpoint
    checkpoint = load_checkpoint()
    if checkpoint:
        print(f"\n⚠ Found checkpoint from {checkpoint['timestamp']}")
        print(f"   Progress: {checkpoint['current_index']}/{checkpoint['total_institutions']} institutions")
        print(f"   Experts collected: {checkpoint['experts_collected']}")

        resume = input("\nResume from checkpoint? (y/n): ").strip().lower()
        if resume == 'y':
            master_list = checkpoint['data']
            start_index = checkpoint['current_index']
            print("✓ Resuming from checkpoint...")

    try:
        # ===== STEP 1: Login =====
        driver.get(LOGIN_URL)
        print("\n" + "-" * 50)
        print("STEP 1: AUTHENTICATION")
        print("-" * 50)
        print("Please log in to the portal manually.")
        print("Navigate to the Experts Dashboard after logging in.")
        input("\n>>> Press ENTER here once you're on the Experts page...")

        # ===== STEP 2: Get Institution List =====
        institution_names = get_institution_names(driver, wait)

        if not institution_names:
            print("\n✗ FATAL: No institutions found. Exiting.")
            return

        # ===== STEP 3: Loop Through All Institutions =====
        print("\n" + "-" * 50)
        print("STEP 2: SCRAPING EXPERTS")
        print("-" * 50)

        for i, uni_name in enumerate(institution_names[start_index:], start=start_index):
            print(f"\n[{i + 1}/{len(institution_names)}] {uni_name}")

            # Select institution from dropdown
            success = select_institution(driver, wait, uni_name)

            if success:
                # Scrape all pages for this institution
                experts = scrape_all_pages_for_institution(driver, wait, uni_name)
                master_list.extend(experts)
                print(f"      ✓ Collected {len(experts)} experts (Running total: {len(master_list)})")
            else:
                print(f"      ✗ Skipped due to selection error")

            # Save checkpoint after each institution
            save_checkpoint(master_list, i + 1, institution_names)

            # Small delay between institutions
            time.sleep(0.5)

        # ===== STEP 4: Save Final Results =====
        print("\n" + "=" * 60)
        print("SCRAPING COMPLETE!")
        print("=" * 60)
        print(f"Total Experts Collected: {len(master_list)}")
        print(f"Total Institutions Processed: {len(institution_names)}")

        # Save JSON
        with open(OUTPUT_JSON, 'w', encoding='utf-8') as f:
            json.dump(master_list, f, indent=2, ensure_ascii=False)
        print(f"\n✓ Saved to {OUTPUT_JSON}")

        # Save Excel
        try:
            df = pd.DataFrame(master_list)
            df.to_excel(OUTPUT_EXCEL, index=False, engine='openpyxl')
            print(f"✓ Saved to {OUTPUT_EXCEL}")
        except Exception as e:
            print(f"✗ Excel save failed: {e}")
            # Fallback to CSV
            try:
                df.to_csv("experts_master_list.csv", index=False)
                print("✓ Saved to experts_master_list.csv (fallback)")
            except:
                pass

        # Summary statistics
        print("\n" + "-" * 50)
        print("SUMMARY BY INSTITUTION:")
        print("-" * 50)
        df = pd.DataFrame(master_list)
        if not df.empty:
            summary = df.groupby('Institution').size().sort_values(ascending=False)
            for inst, count in summary.head(10).items():
                print(f"   {inst}: {count}")
            if len(summary) > 10:
                print(f"   ... and {len(summary) - 10} more institutions")

    except KeyboardInterrupt:
        print("\n\n⚠ Script interrupted by user")
        print(f"   Progress saved. Collected {len(master_list)} experts so far.")

        # Save what we have
        with open(OUTPUT_JSON, 'w', encoding='utf-8') as f:
            json.dump(master_list, f, indent=2, ensure_ascii=False)
        print(f"   Saved partial results to {OUTPUT_JSON}")

    except Exception as e:
        print(f"\n✗ CRITICAL ERROR: {e}")
        import traceback
        traceback.print_exc()

        # Emergency save
        if master_list:
            emergency_file = f"emergency_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
            with open(emergency_file, 'w', encoding='utf-8') as f:
                json.dump(master_list, f, indent=2)
            print(f"   Emergency backup saved to {emergency_file}")

    finally:
        print("\n" + "=" * 60)
        print("Script finished. Browser left open for inspection.")
        print("=" * 60)


if __name__ == "__main__":
    main()