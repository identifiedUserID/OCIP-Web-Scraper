"""
PHASE 5: Organization Metadata Harvester (CORRECTED)
=====================================================
This script scrapes organization metadata from the BusinessAdmin index.
It iterates through the paginated table and captures all fields correctly.

COLUMN MAPPING (based on XPath indices):
- td[4] (index 3) = Organization Name
- td[5] (index 4) = Provinces
- td[6] (index 5) = Sectors
- td[7] (index 6) = Requests?
- td[8] (index 7) = Projects?
- td[9] (index 8) = Enabled
- td[10] (index 9) = Actions (Manage Link)

* Checkpoints are saved after every page to ensure data access during runtime. *

Author: AI Assistant
Version: 2.0 (Corrected column mapping)
"""

import time
import json
import re
import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    StaleElementReferenceException
)
from datetime import datetime

# ==========================================
# CONFIGURATION
# ==========================================
LOGIN_URL = "https://www.ocip.express/"
TARGET_URL = "https://www.ocip.express/BusinessAdmin/Index"
OUTPUT_JSON = "organizations_master_list.json"
OUTPUT_EXCEL = "organizations_master_list.xlsx"
CHECKPOINT_FILE = "organizations_checkpoint.json"

# Timing Configuration
PAGE_LOAD_WAIT = 2.0
PAGINATION_WAIT = 1.5
LOADING_MASK_TIMEOUT = 10

# ==========================================
# COLUMN INDEX MAPPING (0-based indices)
# Based on user-provided XPaths:
# td[4] = index 3, td[5] = index 4, etc.
# ==========================================
COL_ORGANIZATION_NAME = 3  # td[4]
COL_PROVINCES = 4  # td[5]
COL_SECTORS = 5  # td[6]
COL_REQUESTS = 6  # td[7]
COL_PROJECTS = 7  # td[8]
COL_ENABLED = 8  # td[9]
COL_ACTIONS = 9  # td[10]


# ==========================================
# DRIVER SETUP
# ==========================================
def get_driver():
    """Initialize Chrome WebDriver."""
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    options.add_experimental_option("detach", True)

    driver = webdriver.Chrome(options=options)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    return driver


# ==========================================
# UTILITY FUNCTIONS
# ==========================================
def wait_for_loading_complete(driver, timeout=LOADING_MASK_TIMEOUT):
    """Wait for loading masks to disappear."""
    try:
        time.sleep(0.3)
        WebDriverWait(driver, timeout).until(
            EC.invisibility_of_element_located((By.CSS_SELECTOR, ".k-loading-mask"))
        )
    except:
        pass


def parse_pagination_info(driver):
    """Parse pagination text (e.g., '1 - 50 of 120 items')."""
    try:
        pager_info = driver.find_element(By.CSS_SELECTOR, "span.k-pager-info.k-label")
        info_text = pager_info.text.strip()
        match = re.match(r'(\d+)\s*-\s*(\d+)\s+of\s+(\d+)\s+items?', info_text)
        if match:
            return (int(match.group(1)), int(match.group(2)), int(match.group(3)))
        return (0, 0, 0)
    except:
        return (0, 0, 0)


def has_next_page(driver):
    """Check if next page button is active."""
    try:
        next_btn = driver.find_element(By.CSS_SELECTOR, "a.k-pager-nav[aria-label='Go to the next page']")
        return next_btn.get_attribute("aria-disabled") != "true"
    except:
        return False


def click_next_page(driver, wait):
    """Click next page button."""
    try:
        next_btn = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "a.k-pager-nav[aria-label='Go to the next page']")
        ))

        if next_btn.get_attribute("aria-disabled") == "true":
            return False

        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_btn)
        time.sleep(0.2)

        try:
            next_btn.click()
        except:
            driver.execute_script("arguments[0].click();", next_btn)

        wait_for_loading_complete(driver)
        time.sleep(PAGINATION_WAIT)
        return True
    except:
        return False


def save_checkpoint(data, current_page, total_items):
    """
    Save progress immediately to disk.
    File is accessible in real-time during script execution.
    """
    checkpoint = {
        "timestamp": datetime.now().isoformat(),
        "last_page_scraped": current_page,
        "total_items_in_table": total_items,
        "organizations_collected": len(data),
        "data": data
    }
    try:
        with open(CHECKPOINT_FILE, 'w', encoding='utf-8') as f:
            json.dump(checkpoint, f, indent=2, ensure_ascii=False)
            f.flush()
            os.fsync(f.fileno())  # Force write to disk
        print(f"      [Checkpoint saved: {len(data)} records]")
    except Exception as e:
        print(f"      [Warning] Checkpoint save failed: {e}")


def load_checkpoint():
    """Load progress."""
    try:
        with open(CHECKPOINT_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except:
        return None


def clean_text(text):
    """Clean and normalize extracted text."""
    if not text:
        return ""
    text = re.sub(r'\s+', ' ', text.strip())
    text = text.replace('\xa0', ' ')
    return text


def parse_yes_no_cell(cell):
    """
    Parse Yes/No checkbox cells.
    These typically have icons or checkbox inputs.
    """
    try:
        # Check for checkbox icons
        icon = cell.find_element(By.CSS_SELECTOR, "span.k-icon")
        icon_class = icon.get_attribute("class") or ""
        icon_title = icon.get_attribute("title") or ""

        if icon_title:
            return icon_title  # Returns "Yes" or "No"

        if "k-i-checkbox-checked" in icon_class or "k-i-check" in icon_class:
            return "Yes"
        elif "k-i-checkbox" in icon_class or "k-i-close" in icon_class or "k-i-x" in icon_class:
            return "No"
    except NoSuchElementException:
        pass

    try:
        # Check for input checkbox
        checkbox = cell.find_element(By.CSS_SELECTOR, "input[type='checkbox']")
        if checkbox.is_selected() or checkbox.get_attribute("checked"):
            return "Yes"
        else:
            return "No"
    except NoSuchElementException:
        pass

    # Fallback: check text content
    text = clean_text(cell.text).lower()
    if text in ["yes", "true", "1", "✓", "✔"]:
        return "Yes"
    elif text in ["no", "false", "0", "✗", "✘", ""]:
        return "No"

    return clean_text(cell.text) if cell.text else "No"


def get_cell_text_safe(cells, index):
    """Safely get text from a cell at given index."""
    try:
        if index < len(cells):
            return clean_text(cells[index].text)
    except:
        pass
    return ""


def get_cell_element_safe(cells, index):
    """Safely get cell element at given index."""
    try:
        if index < len(cells):
            return cells[index]
    except:
        pass
    return None


# ==========================================
# MODULE: Scrape Organization Table
# ==========================================
def scrape_current_page(driver):
    """
    Scrape organization rows with CORRECT column mapping.

    Table structure (from user XPaths):
    - Row: /html/body/div[2]/div/div[5]/div[3]/div[4]/table/tbody/tr[n]
    - td[4] (index 3) = Organization Name
    - td[5] (index 4) = Provinces
    - td[6] (index 5) = Sectors
    - td[7] (index 6) = Requests?
    - td[8] (index 7) = Projects?
    - td[9] (index 8) = Enabled
    - td[10] (index 9) = Actions (Manage Link)
    """
    try:
        # Find all data rows (k-master-row for Kendo grids)
        rows = driver.find_elements(By.CSS_SELECTOR, "tr.k-master-row")
    except Exception as e:
        print(f"      [Error] Could not find rows: {e}")
        return []

    if not rows:
        print("      [Warning] No rows found on this page")
        return []

    page_data = []

    for row_idx, row in enumerate(rows):
        try:
            cells = row.find_elements(By.TAG_NAME, "td")

            # Verify we have enough cells
            if len(cells) < 10:
                print(f"      [Warning] Row {row_idx + 1} has only {len(cells)} cells, expected at least 10. Skipping.")
                continue

            # ==========================================
            # EXTRACT DATA USING CORRECT INDICES
            # ==========================================

            # Organization Name - td[4] (index 3)
            organization_name = get_cell_text_safe(cells, COL_ORGANIZATION_NAME)

            # Provinces - td[5] (index 4)
            provinces = get_cell_text_safe(cells, COL_PROVINCES)

            # Sectors - td[6] (index 5)
            sectors = get_cell_text_safe(cells, COL_SECTORS)

            # Requests? - td[7] (index 6)
            requests_cell = get_cell_element_safe(cells, COL_REQUESTS)
            requests_flag = parse_yes_no_cell(requests_cell) if requests_cell else "No"

            # Projects? - td[8] (index 7)
            projects_cell = get_cell_element_safe(cells, COL_PROJECTS)
            projects_flag = parse_yes_no_cell(projects_cell) if projects_cell else "No"

            # Enabled - td[9] (index 8)
            enabled_cell = get_cell_element_safe(cells, COL_ENABLED)
            enabled_flag = parse_yes_no_cell(enabled_cell) if enabled_cell else "No"

            # Actions / Manage Link - td[10] (index 9)
            manage_url = "Not Found"
            actions_cell = get_cell_element_safe(cells, COL_ACTIONS)

            if actions_cell:
                try:
                    # Look for the first link in the actions cell
                    # XPath showed: td[10]/div/a[1]
                    link_elem = actions_cell.find_element(By.CSS_SELECTOR, "a")
                    manage_url = link_elem.get_attribute("href") or "Not Found"
                except NoSuchElementException:
                    # Try finding any link
                    try:
                        links = actions_cell.find_elements(By.TAG_NAME, "a")
                        if links:
                            manage_url = links[0].get_attribute("href") or "Not Found"
                    except:
                        pass

            # ==========================================
            # BUILD RECORD
            # ==========================================
            record = {
                "Organization_Name": organization_name,
                "Provinces": provinces,
                "Sectors": sectors,
                "Requests": requests_flag,
                "Projects": projects_flag,
                "Enabled": enabled_flag,
                "Manage_URL": manage_url,
                "Scraped_At": datetime.now().isoformat()
            }

            # Only add if we have a valid organization name
            if organization_name:
                page_data.append(record)
            else:
                print(f"      [Warning] Row {row_idx + 1} has empty organization name. Skipping.")

        except StaleElementReferenceException:
            print(f"      [Warning] Stale element on row {row_idx + 1}. Skipping.")
            continue
        except Exception as e:
            print(f"      [Warning] Error on row {row_idx + 1}: {e}")
            continue

    return page_data


def debug_first_row(driver):
    """
    Debug function to print the contents of each cell in the first row.
    Helps verify column mapping is correct.
    """
    print("\n      [DEBUG] Inspecting first row cells:")
    try:
        rows = driver.find_elements(By.CSS_SELECTOR, "tr.k-master-row")
        if rows:
            first_row = rows[0]
            cells = first_row.find_elements(By.TAG_NAME, "td")
            print(f"      Total cells in row: {len(cells)}")
            for idx, cell in enumerate(cells):
                text = clean_text(cell.text)[:50]  # First 50 chars
                print(f"         td[{idx + 1}] (index {idx}): '{text}'")
    except Exception as e:
        print(f"      [DEBUG ERROR] {e}")
    print()


# ==========================================
# MAIN EXECUTION
# ==========================================
def main():
    print("\n" + "=" * 60)
    print("  PHASE 5: ORGANIZATION METADATA HARVESTER (CORRECTED)")
    print("=" * 60)
    print("\n  Column Mapping:")
    print("    - td[4] (index 3) = Organization Name")
    print("    - td[5] (index 4) = Provinces")
    print("    - td[6] (index 5) = Sectors")
    print("    - td[7] (index 6) = Requests?")
    print("    - td[8] (index 7) = Projects?")
    print("    - td[9] (index 8) = Enabled")
    print("    - td[10] (index 9) = Actions (Manage Link)")
    print()

    driver = get_driver()
    wait = WebDriverWait(driver, 20)
    master_list = []
    current_page = 1

    # Load Checkpoint
    checkpoint = load_checkpoint()
    if checkpoint:
        print(f"\n⚠ Found checkpoint: {checkpoint['organizations_collected']} organizations collected")
        print(f"   Last page processed: {checkpoint['last_page_scraped']}")
        if input("Resume from checkpoint? (y/n): ").lower() == 'y':
            master_list = checkpoint['data']
            print("   ✓ Loaded existing data. Will skip duplicates based on Manage_URL.")

    try:
        # Step 1: Login
        driver.get(LOGIN_URL)
        print("\n" + "-" * 50)
        print("STEP 1: AUTHENTICATION")
        print("-" * 50)
        print("Please log in to the portal manually.")
        input(">>> Press ENTER when logged in...")

        # Step 2: Navigate to Organizations
        driver.get(TARGET_URL)
        print(f"\nNavigating to: {TARGET_URL}")
        wait_for_loading_complete(driver)
        time.sleep(PAGE_LOAD_WAIT)

        # Step 3: Get total count and debug first row
        print("\n" + "-" * 50)
        print("STEP 2: SCRAPING ORGANIZATIONS")
        print("-" * 50)

        start, end, total = parse_pagination_info(driver)
        print(f"\n      Total items in table: {total}")

        # Debug: show first row to verify mapping
        debug_first_row(driver)

        # Ask user to confirm mapping looks correct
        proceed = input("Does the column mapping look correct? (y/n): ").strip().lower()
        if proceed != 'y':
            print("Exiting. Please check the column indices and update the script.")
            return

        # Build set of existing URLs for deduplication
        existing_urls = {item['Manage_URL'] for item in master_list}

        # Step 4: Pagination Loop
        print("\n      Starting scrape...")

        while True:
            # Scrape current page
            page_data = scrape_current_page(driver)

            # Add new records (skip duplicates)
            new_count = 0
            for item in page_data:
                if item['Manage_URL'] not in existing_urls:
                    master_list.append(item)
                    existing_urls.add(item['Manage_URL'])
                    new_count += 1

            print(f"\n      Page {current_page}: Found {len(page_data)} rows, {new_count} new records")
            print(f"      Running total: {len(master_list)} organizations")

            # Sample output for verification
            if page_data and current_page == 1:
                print("\n      [Sample] First record on this page:")
                sample = page_data[0]
                for key, value in sample.items():
                    if key != "Scraped_At":
                        print(f"         {key}: {value}")

            # SAVE CHECKPOINT IMMEDIATELY
            save_checkpoint(master_list, current_page, total)

            # Pagination Check
            start, end, total = parse_pagination_info(driver)
            if end >= total:
                print(f"\n      → Reached last page (showing {start}-{end} of {total})")
                break

            if not has_next_page(driver):
                print("\n      → No more pages available")
                break

            # Next Page
            if click_next_page(driver, wait):
                current_page += 1
            else:
                print("\n      → Could not click next page")
                break

            # Safety break
            if current_page > 500:
                print("\n      → Safety limit reached (500 pages)")
                break

        # Step 5: Save Final Results
        print("\n" + "=" * 60)
        print("SCRAPING COMPLETE!")
        print("=" * 60)
        print(f"\nTotal Organizations Collected: {len(master_list)}")
        print(f"Pages Processed: {current_page}")

        # Save JSON
        with open(OUTPUT_JSON, 'w', encoding='utf-8') as f:
            json.dump(master_list, f, indent=2, ensure_ascii=False)
        print(f"\n✓ Saved to {OUTPUT_JSON}")

        # Save Excel
        try:
            df = pd.DataFrame(master_list)
            # Reorder columns for clarity
            column_order = [
                "Organization_Name",
                "Provinces",
                "Sectors",
                "Requests",
                "Projects",
                "Enabled",
                "Manage_URL",
                "Scraped_At"
            ]
            df = df[[col for col in column_order if col in df.columns]]
            df.to_excel(OUTPUT_EXCEL, index=False, engine='openpyxl')
            print(f"✓ Saved to {OUTPUT_EXCEL}")
        except Exception as e:
            print(f"✗ Excel save failed: {e}")
            # Fallback to CSV
            try:
                df.to_csv("organizations_master_list.csv", index=False)
                print("✓ Saved to organizations_master_list.csv (fallback)")
            except:
                pass

        # Summary Statistics
        print("\n" + "-" * 50)
        print("SUMMARY:")
        print("-" * 50)

        df = pd.DataFrame(master_list)
        if not df.empty:
            print(f"   Total Organizations: {len(df)}")

            # Count by flags
            if 'Requests' in df.columns:
                requests_yes = (df['Requests'] == 'Yes').sum()
                print(f"   Organizations with Requests: {requests_yes}")

            if 'Projects' in df.columns:
                projects_yes = (df['Projects'] == 'Yes').sum()
                print(f"   Organizations with Projects: {projects_yes}")

            if 'Enabled' in df.columns:
                enabled_yes = (df['Enabled'] == 'Yes').sum()
                print(f"   Enabled Organizations: {enabled_yes}")

            # Count by province (if available)
            if 'Provinces' in df.columns:
                provinces_count = df['Provinces'].value_counts().head(5)
                print(f"\n   Top Provinces:")
                for prov, count in provinces_count.items():
                    if prov:
                        print(f"      {prov}: {count}")

    except KeyboardInterrupt:
        print("\n\n⚠ Script interrupted by user")
        save_checkpoint(master_list, current_page, total if 'total' in dir() else 0)
        print(f"   Progress saved. Collected {len(master_list)} organizations.")

    except Exception as e:
        print(f"\n✗ CRITICAL ERROR: {e}")
        import traceback
        traceback.print_exc()

        # Emergency save
        if master_list:
            emergency_file = f"emergency_phase5_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
            with open(emergency_file, 'w', encoding='utf-8') as f:
                json.dump(master_list, f, indent=2)
            print(f"   Emergency backup saved to {emergency_file}")

    finally:
        print("\n" + "=" * 60)
        print("Script finished. Browser left open for inspection.")
        print(f"Checkpoint file: {CHECKPOINT_FILE}")
        print("=" * 60)


if __name__ == "__main__":
    main()