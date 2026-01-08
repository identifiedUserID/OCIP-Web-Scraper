"""
PHASE 3: Facility Harvester Script
==================================
This script scrapes facility metadata from the eCampusOntario OCIP Express portal.
It navigates the HEI dropdown, handles pagination, and extracts the 'Manage' links
from the specific table column (16th column) identified.

Author: AI Assistant
Version: 1.0
Based on Phase 1 Framework
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
TARGET_URL = "https://www.ocip.express/FacilityAdmin/Index"
OUTPUT_JSON = "facilities_master_list.json"
OUTPUT_EXCEL = "facilities_master_list.xlsx"
CHECKPOINT_FILE = "phase3_checkpoint.json"

# Timing Configuration
PAGE_LOAD_WAIT = 2.0
PAGINATION_WAIT = 1.5
DROPDOWN_CLOSE_WAIT = 0.5
LOADING_MASK_TIMEOUT = 10

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

def reset_to_first_page(driver, wait):
    """Reset grid to page 1."""
    try:
        start, end, total = parse_pagination_info(driver)
        if start <= 1: return True

        first_btn = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "a.k-pager-nav.k-pager-first[aria-label='Go to the first page']")
        ))
        if first_btn.get_attribute("aria-disabled") == "true": return True

        driver.execute_script("arguments[0].click();", first_btn)
        wait_for_loading_complete(driver)
        time.sleep(PAGINATION_WAIT)
        print("         → Reset to page 1")
        return True
    except:
        return False

def save_checkpoint(data, current_index, institution_names):
    """Save progress."""
    checkpoint = {
        "timestamp": datetime.now().isoformat(),
        "current_index": current_index,
        "total_institutions": len(institution_names),
        "facilities_collected": len(data),
        "data": data
    }
    with open(CHECKPOINT_FILE, 'w', encoding='utf-8') as f:
        json.dump(checkpoint, f, indent=2)

def load_checkpoint():
    """Load progress."""
    try:
        with open(CHECKPOINT_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except:
        return None

# ==========================================
# MODULE 1: Get Institution Names
# ==========================================
def get_institution_names(driver, wait):
    """Get list of institutions from the filter dropdown."""
    print("\n" + "=" * 50)
    print("GATHERING INSTITUTION LIST")
    print("=" * 50)
    try:
        # Targeting the dropdown arrow/container
        # Based on user input: /html/body/div[2]/div/div[5]/div[3]/div[2]/div[1]/div/div[2]/span/span[1]
        # We look for the aria-controls usually associated with 'HeiId_listbox' or similar
        
        dropdown_trigger = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "span[aria-controls$='listbox']") # Generic 'ends with listbox' to be safe
        ))
        dropdown_trigger.click()

        # Wait for listbox (Checking standard HeiId_listbox or just any visible listbox)
        listbox = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "ul[id$='listbox'][aria-hidden='false']")))
        time.sleep(0.5)

        options = listbox.find_elements(By.TAG_NAME, "li")
        names = []
        for opt in options:
            text = opt.text.strip()
            if text and "Select HEI" not in text and text != "":
                names.append(text)

        print(f"✓ Found {len(names)} institutions to process")
        
        # Close dropdown
        driver.find_element(By.TAG_NAME, "body").click()
        time.sleep(DROPDOWN_CLOSE_WAIT)
        return names

    except Exception as e:
        print(f"✗ FAILED to get institutions: {e}")
        return []

# ==========================================
# MODULE 2: Select Institution
# ==========================================
def select_institution(driver, wait, target_name):
    """Select specific institution from dropdown."""
    try:
        dropdown_trigger = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "span[aria-controls$='listbox']")
        ))
        dropdown_trigger.click()

        listbox = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "ul[id$='listbox'][aria-hidden='false']")))
        time.sleep(0.3)

        all_options = listbox.find_elements(By.TAG_NAME, "li")
        target_element = None
        for opt in all_options:
            if target_name in opt.text:
                target_element = opt
                break

        if target_element:
            driver.execute_script("arguments[0].click();", target_element)
            wait_for_loading_complete(driver)
            time.sleep(PAGE_LOAD_WAIT)
            return True
        else:
            driver.find_element(By.TAG_NAME, "body").click()
            return False
    except:
        try:
            driver.find_element(By.TAG_NAME, "body").click()
        except:
            pass
        return False

# ==========================================
# MODULE 3: Scrape Facility Table
# ==========================================
def scrape_current_page(driver, institution_name):
    """
    Scrape facility rows.
    Targeting specific table structure where Manage link is in the 16th column (index 15).
    """
    try:
        # Table container: /html/body/div[2]/div/div[5]/div[4]
        # Rows: .k-master-row
        rows = driver.find_elements(By.CSS_SELECTOR, "tr.k-master-row")
    except:
        return []

    if not rows:
        return []

    page_data = []
    
    for row in rows:
        try:
            cells = row.find_elements(By.TAG_NAME, "td")
            
            # User specified logic: tr[1]/td[16]/div
            # Index 15 corresponds to the 16th td
            if len(cells) < 16:
                # Fallback if table is smaller than expected
                continue

            # Extract basic visible data (indices are approximate, assuming standard layout)
            # Usually: ID(0 or 1), Name(2 or 3), Type(4), etc.
            # We grab a few early columns to ensure we have identification
            try:
                facility_id = cells[1].text.strip() # Guessing ID column
                facility_name = cells[3].text.strip() # Guessing Name column
                facility_type = cells[4].text.strip() # Guessing Type column
            except:
                facility_id = "Unknown"
                facility_name = cells[2].text.strip() if len(cells) > 2 else "Unknown"
                facility_type = ""

            # --- EXTRACT MANAGE LINK (CRITICAL STEP) ---
            manage_url = "Not Found"
            try:
                # Target the 16th cell (index 15)
                target_cell = cells[15] 
                
                # Look for anchor tag inside the div inside the cell
                link_elem = target_cell.find_element(By.TAG_NAME, "a")
                manage_url = link_elem.get_attribute("href")
                
                # Double check title just in case
                if not manage_url and "View" not in link_elem.get_attribute("title"):
                     # Try finding any link in that cell
                     pass
            except NoSuchElementException:
                # Fallback: Look for any link with "Details" or "View" in the whole row
                try:
                    links = row.find_elements(By.CSS_SELECTOR, "a[href*='Details'], a[title='View Full Details']")
                    if links:
                        manage_url = links[0].get_attribute("href")
                except:
                    pass

            record = {
                "Institution": institution_name,
                "Facility_Name": facility_name,
                "Facility_ID": facility_id,
                "Type": facility_type,
                "Manage_URL": manage_url,
                "Scraped_At": datetime.now().isoformat()
            }
            
            # Only add if we found a URL (or decide to keep all rows)
            page_data.append(record)

        except StaleElementReferenceException:
            continue
        except Exception as e:
            continue

    return page_data

# ==========================================
# MODULE 4: Pagination Loop
# ==========================================
def scrape_all_pages_for_institution(driver, wait, institution_name):
    """Scrape all pages for a facility."""
    all_facilities = []
    current_page = 1
    
    start, end, total = parse_pagination_info(driver)
    
    if total == 0:
        print(f"      → No facilities found")
        return []

    print(f"      → Found {total} facilities")

    while True:
        page_data = scrape_current_page(driver, institution_name)
        all_facilities.extend(page_data)
        print(f"         Page {current_page}: Scraped {len(page_data)} records")

        start, end, total = parse_pagination_info(driver)
        if end >= total or not has_next_page(driver):
            break

        if click_next_page(driver, wait):
            current_page += 1
        else:
            break
            
        if current_page > 50: break # Safety

    if current_page > 1:
        reset_to_first_page(driver, wait)

    return all_facilities

# ==========================================
# MAIN EXECUTION
# ==========================================
def main():
    print("\n" + "=" * 60)
    print("  PHASE 3: FACILITY METADATA HARVESTER")
    print("=" * 60)

    driver = get_driver()
    wait = WebDriverWait(driver, 20)
    master_list = []
    start_index = 0

    # Load Checkpoint
    checkpoint = load_checkpoint()
    if checkpoint:
        print(f"\n⚠ Found checkpoint: {checkpoint['facilities_collected']} collected")
        if input("Resume? (y/n): ").lower() == 'y':
            master_list = checkpoint['data']
            start_index = checkpoint['current_index']

    try:
        # Step 1: Login
        driver.get(LOGIN_URL)
        print("\nSTEP 1: Please Log In.")
        input(">>> Press ENTER when logged in...")

        # Step 2: Navigate to Facilities
        driver.get(TARGET_URL)
        print(f"\nNavigating to: {TARGET_URL}")
        time.sleep(PAGE_LOAD_WAIT)

        # Step 3: Get Institutions
        institution_names = get_institution_names(driver, wait)
        if not institution_names:
            print("✗ No institutions found in dropdown.")
            return

        # Step 4: Loop
        print("\nSTEP 2: SCRAPING FACILITIES")
        for i, uni_name in enumerate(institution_names[start_index:], start=start_index):
            print(f"\n[{i + 1}/{len(institution_names)}] {uni_name}")

            if select_institution(driver, wait, uni_name):
                data = scrape_all_pages_for_institution(driver, wait, uni_name)
                master_list.extend(data)
                print(f"      ✓ Total collected: {len(master_list)}")
            else:
                print("      ✗ Selection failed")

            save_checkpoint(master_list, i + 1, institution_names)
            time.sleep(0.5)

        # Step 5: Save
        print("\n" + "=" * 60)
        print("SAVING RESULTS")
        
        with open(OUTPUT_JSON, 'w', encoding='utf-8') as f:
            json.dump(master_list, f, indent=2)
            
        try:
            pd.DataFrame(master_list).to_excel(OUTPUT_EXCEL, index=False)
            print(f"✓ Saved to {OUTPUT_EXCEL}")
        except:
            pass
            
    except Exception as e:
        print(f"\n✗ CRITICAL ERROR: {e}")
        import traceback
        traceback.print_exc()
        
    finally:
        print("\nScript finished.")

if __name__ == "__main__":
    main()