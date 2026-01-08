"""
PHASE 2: Expert Deep Scraper
============================
This script visits each expert's detail page (from Phase 1 master list)
and extracts all information from every accordion section.

Author: AI Assistant
Version: 1.0
"""

import time
import json
import re
import os
from datetime import datetime
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

# ==========================================
# CONFIGURATION
# ==========================================
INPUT_FILE = "experts_master_list.json"
OUTPUT_FILE = "experts_full_details.json"
CHECKPOINT_FILE = "phase2_checkpoint.json"
ERROR_LOG_FILE = "phase2_errors.json"

# Timing Configuration
PAGE_LOAD_WAIT = 2.0
ACCORDION_EXPAND_WAIT = 0.5
BETWEEN_EXPERTS_DELAY = 1.0
REQUEST_DELAY = 0.3  # Delay between accordion clicks

# Rate limiting - pause every N experts
BATCH_SIZE = 50
BATCH_PAUSE = 10  # seconds


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
def load_master_list(filepath):
    """Load the master list from Phase 1."""
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            data = json.load(f)
        print(f"✓ Loaded {len(data)} experts from {filepath}")
        return data
    except FileNotFoundError:
        print(f"✗ ERROR: {filepath} not found. Run Phase 1 first.")
        return None
    except json.JSONDecodeError as e:
        print(f"✗ ERROR: Invalid JSON in {filepath}: {e}")
        return None


def save_checkpoint(processed_data, current_index, errors):
    """Save progress checkpoint."""
    checkpoint = {
        "timestamp": datetime.now().isoformat(),
        "current_index": current_index,
        "experts_processed": len(processed_data),
        "errors_count": len(errors),
        "data": processed_data
    }
    with open(CHECKPOINT_FILE, 'w', encoding='utf-8') as f:
        json.dump(checkpoint, f, indent=2, ensure_ascii=False)

    # Save errors separately
    with open(ERROR_LOG_FILE, 'w', encoding='utf-8') as f:
        json.dump(errors, f, indent=2, ensure_ascii=False)


def load_checkpoint():
    """Load previous checkpoint if exists."""
    try:
        with open(CHECKPOINT_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return None


def clean_text(text):
    """Clean and normalize extracted text."""
    if not text:
        return ""
    # Remove excess whitespace
    text = re.sub(r'\s+', ' ', text.strip())
    # Remove non-breaking spaces
    text = text.replace('\xa0', ' ')
    return text


def parse_yes_no(element):
    """Parse Yes/No icon elements."""
    try:
        icon = element.find_element(By.CSS_SELECTOR, "span.k-icon")
        title = icon.get_attribute("title")
        if title:
            return title  # Returns "Yes" or "No"
        # Fallback: check class names
        classes = icon.get_attribute("class")
        if "k-i-checkbox-checked" in classes:
            return "Yes"
        elif "k-i-checkbox" in classes:
            return "No"
    except:
        pass
    return clean_text(element.text)


def expand_all_accordions(driver, wait):
    """Expand all collapsed accordion panels."""
    try:
        # Find all collapsed accordion headers
        collapsed_headers = driver.find_elements(
            By.CSS_SELECTOR,
            "li.k-panelbar-header[aria-expanded='false'] > a.k-link"
        )

        for header in collapsed_headers:
            try:
                # Scroll into view
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", header)
                time.sleep(0.1)

                # Click to expand
                try:
                    header.click()
                except ElementClickInterceptedException:
                    driver.execute_script("arguments[0].click();", header)

                time.sleep(ACCORDION_EXPAND_WAIT)

            except StaleElementReferenceException:
                continue
            except Exception as e:
                continue

        # Give time for all content to render
        time.sleep(0.5)
        return True

    except Exception as e:
        print(f"      [Warning] Accordion expansion error: {e}")
        return False


# ==========================================
# SECTION EXTRACTORS
# ==========================================

def extract_general_information(driver):
    """Extract data from General Information section."""
    data = {}

    try:
        # Find the General Information panel
        panel = driver.find_element(By.ID, "ProfileBar-1")

        # Extract Academic Unit breadcrumb
        try:
            breadcrumb_items = panel.find_elements(By.CSS_SELECTOR, "ol.breadcrumb li")
            academic_unit_path = []
            for item in breadcrumb_items:
                text = clean_text(item.text)
                if text:
                    academic_unit_path.append(text)
            data["Academic_Unit"] = " > ".join(academic_unit_path) if academic_unit_path else ""
        except:
            data["Academic_Unit"] = ""

        # Extract key-value pairs from rows
        rows = panel.find_elements(By.CSS_SELECTOR, "div.row")

        for row in rows:
            try:
                labels = row.find_elements(By.TAG_NAME, "label")
                if not labels:
                    continue

                label = labels[0]
                label_for = label.get_attribute("for") or ""
                label_text = clean_text(label.text)

                # Find the corresponding value column
                cols = row.find_elements(By.CSS_SELECTOR, "div[class*='col-md']")
                value_col = None

                for col in cols:
                    if label not in col.find_elements(By.TAG_NAME, "label"):
                        # Check if this column has the value
                        if col.text.strip() or col.find_elements(By.CSS_SELECTOR, "span.k-icon, a"):
                            value_col = col
                            break

                if not value_col:
                    continue

                # Extract value based on field type
                field_key = label_for if label_for else label_text.replace(" ", "_")

                if label_for in ["IsLinkedToUser", "Enabled"]:
                    data[field_key] = parse_yes_no(value_col)
                elif label_for == "Contact":
                    # Extract email and phone
                    try:
                        email_elem = value_col.find_element(By.CSS_SELECTOR, "a[href^='mailto:']")
                        data["Email"] = email_elem.text.strip()
                    except:
                        data["Email"] = ""
                    try:
                        phone_elem = value_col.find_element(By.CSS_SELECTOR, "a[href^='tel:']")
                        data["Phone"] = phone_elem.text.strip()
                    except:
                        data["Phone"] = ""
                elif label_for == "ReputationScore":
                    # Extract rating value
                    try:
                        rating_span = value_col.find_element(By.CSS_SELECTOR, "span.k-rating")
                        rating_value = rating_span.get_attribute("aria-valuenow")
                        data["Reputation_Score"] = rating_value if rating_value else "Not Rated"
                        # Also get the text description
                        rating_text = clean_text(value_col.text)
                        if "Not Rated" in rating_text:
                            data["Reputation_Score"] = "Not Rated"
                    except:
                        data["Reputation_Score"] = clean_text(value_col.text)
                else:
                    data[field_key] = clean_text(value_col.text)

            except Exception as e:
                continue

        # Extract photo URL if present
        try:
            img = panel.find_element(By.CSS_SELECTOR, "img[alt]")
            data["Photo_URL"] = img.get_attribute("src")
        except:
            data["Photo_URL"] = ""

    except NoSuchElementException:
        pass
    except Exception as e:
        print(f"      [Warning] General Information extraction error: {e}")

    return data


def extract_details(driver):
    """Extract data from Details section."""
    data = {}

    try:
        panel = driver.find_element(By.ID, "ProfileBar-2")
        rows = panel.find_elements(By.CSS_SELECTOR, "div.row")

        for row in rows:
            try:
                labels = row.find_elements(By.TAG_NAME, "label")
                if not labels:
                    continue

                label = labels[0]
                label_for = label.get_attribute("for") or ""
                label_text = clean_text(label.text)

                # Find value column
                cols = row.find_elements(By.CSS_SELECTOR, "div.col-md-9, div.col-md-7")
                if cols:
                    value = clean_text(cols[0].text)
                    field_key = label_for if label_for else label_text.replace(" ", "_")
                    data[field_key] = value

            except Exception:
                continue

    except NoSuchElementException:
        pass
    except Exception as e:
        print(f"      [Warning] Details extraction error: {e}")

    return data


def extract_expert_demographics(driver):
    """Extract data from Expert Demographics section."""
    data = {}

    try:
        panel = driver.find_element(By.ID, "ProfileBar-3")
        rows = panel.find_elements(By.CSS_SELECTOR, "div.row")

        for row in rows:
            try:
                labels = row.find_elements(By.TAG_NAME, "label")
                if not labels:
                    continue

                label = labels[0]
                label_text = clean_text(label.text)

                cols = row.find_elements(By.CSS_SELECTOR, "div.col-md-7, div.col-md-9")
                if cols:
                    value = clean_text(cols[0].text)
                    data[label_text.replace(" ", "_")] = value

            except Exception:
                continue

    except NoSuchElementException:
        pass
    except Exception as e:
        print(f"      [Warning] Demographics extraction error: {e}")

    return data


def extract_expertise(driver):
    """Extract data from Expertise section (table/grid)."""
    expertise_list = []

    try:
        panel = driver.find_element(By.ID, "ProfileBar-4")

        # Check for "No records found"
        try:
            no_data = panel.find_element(By.CSS_SELECTOR, "tr.k-no-data, div.k-grid-norecords-template")
            if no_data:
                return []
        except NoSuchElementException:
            pass

        # Find the grid table
        grid = panel.find_element(By.ID, "contactsGrid")
        rows = grid.find_elements(By.CSS_SELECTOR, "tbody tr.k-master-row")

        for row in rows:
            try:
                cells = row.find_elements(By.TAG_NAME, "td")
                if len(cells) >= 4:
                    expertise_entry = {
                        "SRED_Code": clean_text(cells[0].text),
                        "Area": clean_text(cells[1].text),
                        "Discipline": clean_text(cells[2].text),
                        "Field": clean_text(cells[3].text)
                    }
                    expertise_list.append(expertise_entry)
            except Exception:
                continue

    except NoSuchElementException:
        pass
    except Exception as e:
        print(f"      [Warning] Expertise extraction error: {e}")

    return expertise_list


def extract_price_availability(driver):
    """Extract data from Price & Availability section."""
    data = {}

    try:
        panel = driver.find_element(By.ID, "ProfileBar-5")

        # Extract Daily Rate
        try:
            rows = panel.find_elements(By.CSS_SELECTOR, "div.row")
            for row in rows:
                labels = row.find_elements(By.TAG_NAME, "label")
                if labels:
                    label_for = labels[0].get_attribute("for") or ""
                    if label_for == "PerDiemRate":
                        cols = row.find_elements(By.CSS_SELECTOR, "div.col-md-9")
                        if cols:
                            data["Daily_Rate"] = clean_text(cols[0].text)
        except:
            data["Daily_Rate"] = ""

        # Extract availability flags from table
        try:
            table = panel.find_element(By.CSS_SELECTOR, "table.table")
            headers = table.find_elements(By.CSS_SELECTOR, "thead th")
            cells = table.find_elements(By.CSS_SELECTOR, "tbody td")

            header_texts = [
                "Can_Initiate_Innovation_Challenge",
                "Available_for_Scoping",
                "Available_for_Projects",
                "Can_be_Principal_Investigator"
            ]

            for i, cell in enumerate(cells):
                if i < len(header_texts):
                    data[header_texts[i]] = parse_yes_no(cell)

        except Exception as e:
            pass

    except NoSuchElementException:
        pass
    except Exception as e:
        print(f"      [Warning] Price & Availability extraction error: {e}")

    return data


def extract_facility_affiliation(driver):
    """Extract data from Facility Affiliation section (table/grid)."""
    facilities_list = []

    try:
        panel = driver.find_element(By.ID, "ProfileBar-6")

        # Check for "No records found"
        try:
            no_data = panel.find_element(By.CSS_SELECTOR, "tr.k-no-data, div.k-grid-norecords-template")
            if no_data:
                return []
        except NoSuchElementException:
            pass

        grid = panel.find_element(By.ID, "networksGrid")
        rows = grid.find_elements(By.CSS_SELECTOR, "tbody tr.k-master-row")

        for row in rows:
            try:
                cells = row.find_elements(By.TAG_NAME, "td")
                if len(cells) >= 2:
                    facility_entry = {
                        "Facility_Name": clean_text(cells[0].text),
                        "Is_Primary_Facility": parse_yes_no(cells[1])
                    }
                    facilities_list.append(facility_entry)
            except Exception:
                continue

    except NoSuchElementException:
        pass
    except Exception as e:
        print(f"      [Warning] Facility Affiliation extraction error: {e}")

    return facilities_list


def extract_web_presence(driver):
    """Extract data from Web Presence section (table/grid)."""
    web_presence_list = []

    try:
        panel = driver.find_element(By.ID, "ProfileBar-8")

        # Check for "No records found"
        try:
            no_data = panel.find_element(By.CSS_SELECTOR, "tr.k-no-data, div.k-grid-norecords-template")
            if no_data:
                return []
        except NoSuchElementException:
            pass

        # Find the webGrid table
        grid = panel.find_element(By.ID, "webGrid")
        rows = grid.find_elements(By.CSS_SELECTOR, "tbody tr.k-master-row")

        for row in rows:
            try:
                cells = row.find_elements(By.TAG_NAME, "td")
                if len(cells) >= 3:
                    # URL might be in an anchor tag
                    url = ""
                    try:
                        link = cells[2].find_element(By.TAG_NAME, "a")
                        url = link.get_attribute("href")
                    except:
                        url = clean_text(cells[2].text)

                    web_entry = {
                        "Name": clean_text(cells[0].text),
                        "Type": clean_text(cells[1].text),
                        "URL": url
                    }
                    web_presence_list.append(web_entry)
            except Exception:
                continue

    except NoSuchElementException:
        pass
    except Exception as e:
        print(f"      [Warning] Web Presence extraction error: {e}")

    return web_presence_list


def extract_ocip_activity(driver):
    """Extract data from OCIP Activity section (table/grid)."""
    activity_list = []

    try:
        panel = driver.find_element(By.ID, "ProfileBar-9")

        # Check for "No records found"
        try:
            no_data = panel.find_element(By.CSS_SELECTOR, "tr.k-no-data, div.k-grid-norecords-template")
            if no_data:
                return []
        except NoSuchElementException:
            pass

        # Note: This panel also uses id="webGrid" (duplicate ID in HTML)
        # We need to find it within the panel context
        rows = panel.find_elements(By.CSS_SELECTOR, "tbody tr.k-master-row")

        for row in rows:
            try:
                cells = row.find_elements(By.TAG_NAME, "td")
                if len(cells) >= 4:
                    # Project name might be in a link
                    project_name = ""
                    project_url = ""
                    try:
                        link = cells[0].find_element(By.TAG_NAME, "a")
                        project_name = clean_text(link.text)
                        project_url = link.get_attribute("href")
                    except:
                        project_name = clean_text(cells[0].text)

                    activity_entry = {
                        "Project_Name": project_name,
                        "Project_URL": project_url,
                        "Type": clean_text(cells[1].text),
                        "Organization": clean_text(cells[2].text),
                        "Current_Status": clean_text(cells[3].text)
                    }
                    activity_list.append(activity_entry)
            except Exception:
                continue

    except NoSuchElementException:
        pass
    except Exception as e:
        print(f"      [Warning] OCIP Activity extraction error: {e}")

    return activity_list


def extract_audit_trail(driver):
    """Extract data from Audit Trail section."""
    data = {}

    try:
        panel = driver.find_element(By.ID, "ProfileBar-10")
        rows = panel.find_elements(By.CSS_SELECTOR, "div.row")

        for row in rows:
            try:
                labels = row.find_elements(By.TAG_NAME, "label")
                if not labels:
                    continue

                label = labels[0]
                label_for = label.get_attribute("for") or ""
                label_text = clean_text(label.text)

                cols = row.find_elements(By.CSS_SELECTOR, "div.col-md-9")
                if cols:
                    value = clean_text(cols[0].text)
                    field_key = label_for if label_for else label_text.replace(" ", "_")
                    data[field_key] = value

            except Exception:
                continue

    except NoSuchElementException:
        pass
    except Exception as e:
        print(f"      [Warning] Audit Trail extraction error: {e}")

    return data


# ==========================================
# MAIN EXTRACTION FUNCTION
# ==========================================

def extract_expert_full_profile(driver, wait, expert_basic_info):
    """
    Extract all information from an expert's detail page.
    Returns a complete profile dictionary.
    """
    url = expert_basic_info.get("Manage_URL", "")

    if not url or url == "Not Found":
        return None

    try:
        # Navigate to the detail page
        driver.get(url)
        time.sleep(PAGE_LOAD_WAIT)

        # Wait for page to load (look for the panel bar)
        try:
            wait.until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, "ul.k-panelbar, li.k-panelbar-header")
            ))
        except TimeoutException:
            print(f"      [Warning] Page load timeout for {url}")
            return None

        # Expand all accordion sections
        expand_all_accordions(driver, wait)

        # Initialize profile with basic info from Phase 1
        profile = {
            "Meta": {
                "Source_URL": url,
                "Scraped_At": datetime.now().isoformat(),
                "Institution": expert_basic_info.get("Institution", ""),
                "Name_From_List": expert_basic_info.get("Name", ""),
                "Expert_ID": expert_basic_info.get("Expert_ID", ""),
                "Profile_URL": expert_basic_info.get("Profile_URL", "")
            }
        }

        # Extract each section
        profile["General_Information"] = extract_general_information(driver)
        time.sleep(REQUEST_DELAY)

        profile["Details"] = extract_details(driver)
        time.sleep(REQUEST_DELAY)

        profile["Expert_Demographics"] = extract_expert_demographics(driver)
        time.sleep(REQUEST_DELAY)

        profile["Expertise"] = extract_expertise(driver)
        time.sleep(REQUEST_DELAY)

        profile["Price_Availability"] = extract_price_availability(driver)
        time.sleep(REQUEST_DELAY)

        profile["Facility_Affiliation"] = extract_facility_affiliation(driver)
        time.sleep(REQUEST_DELAY)

        profile["Web_Presence"] = extract_web_presence(driver)
        time.sleep(REQUEST_DELAY)

        profile["OCIP_Activity"] = extract_ocip_activity(driver)
        time.sleep(REQUEST_DELAY)

        profile["Audit_Trail"] = extract_audit_trail(driver)

        return profile

    except Exception as e:
        print(f"      [Error] Failed to extract {url}: {e}")
        return None


# ==========================================
# MAIN EXECUTION
# ==========================================

def main():
    """Main execution function for Phase 2."""

    print("\n" + "=" * 60)
    print("  PHASE 2: EXPERT DEEP SCRAPER")
    print("  Extracting Full Profiles from Detail Pages")
    print("=" * 60)

    # Load master list from Phase 1
    master_list = load_master_list(INPUT_FILE)
    if not master_list:
        return

    # Initialize
    driver = get_driver()
    wait = WebDriverWait(driver, 20)
    processed_data = []
    errors = []
    start_index = 0

    # Check for existing checkpoint
    checkpoint = load_checkpoint()
    if checkpoint:
        print(f"\n⚠ Found checkpoint from {checkpoint['timestamp']}")
        print(f"   Progress: {checkpoint['current_index']}/{len(master_list)} experts")
        print(f"   Successfully processed: {checkpoint['experts_processed']}")
        print(f"   Errors: {checkpoint['errors_count']}")

        resume = input("\nResume from checkpoint? (y/n): ").strip().lower()
        if resume == 'y':
            processed_data = checkpoint['data']
            start_index = checkpoint['current_index']
            # Load errors
            try:
                with open(ERROR_LOG_FILE, 'r') as f:
                    errors = json.load(f)
            except:
                errors = []
            print("✓ Resuming from checkpoint...")

    try:
        # Login step
        print("\n" + "-" * 50)
        print("STEP 1: AUTHENTICATION")
        print("-" * 50)
        driver.get("https://www.ocip.express/")
        print("Please log in to the portal manually.")
        input("\n>>> Press ENTER here once you're logged in...")

        # Process each expert
        print("\n" + "-" * 50)
        print("STEP 2: EXTRACTING EXPERT PROFILES")
        print("-" * 50)

        total = len(master_list)

        for i, expert in enumerate(master_list[start_index:], start=start_index):
            expert_name = expert.get("Name", "Unknown")
            institution = expert.get("Institution", "Unknown")
            url = expert.get("Manage_URL", "")

            print(f"\n[{i + 1}/{total}] {expert_name} ({institution})")

            if not url or url == "Not Found":
                print("      → Skipped: No valid URL")
                errors.append({
                    "index": i,
                    "name": expert_name,
                    "reason": "No valid Manage URL"
                })
                continue

            # Extract full profile
            profile = extract_expert_full_profile(driver, wait, expert)

            if profile:
                processed_data.append(profile)
                print(f"      ✓ Extracted successfully")

                # Show summary of what was found
                expertise_count = len(profile.get("Expertise", []))
                facilities_count = len(profile.get("Facility_Affiliation", []))
                has_bio = bool(profile.get("Details", {}).get("ProfileDescription", ""))
                print(
                    f"         Bio: {'Yes' if has_bio else 'No'} | Expertise: {expertise_count} | Facilities: {facilities_count}")
            else:
                errors.append({
                    "index": i,
                    "name": expert_name,
                    "url": url,
                    "reason": "Extraction failed"
                })
                print(f"      ✗ Extraction failed")

            # Save checkpoint periodically
            if (i + 1) % 10 == 0:
                save_checkpoint(processed_data, i + 1, errors)
                print(f"\n   [Checkpoint saved: {len(processed_data)} profiles]")

            # Rate limiting: pause every BATCH_SIZE experts
            if (i + 1) % BATCH_SIZE == 0 and i + 1 < total:
                print(f"\n   [Rate limit pause: {BATCH_PAUSE}s...]")
                time.sleep(BATCH_PAUSE)

            # Delay between experts
            time.sleep(BETWEEN_EXPERTS_DELAY)

        # ===== FINAL SAVE =====
        print("\n" + "=" * 60)
        print("SCRAPING COMPLETE!")
        print("=" * 60)
        print(f"Total Experts Processed: {len(processed_data)}")
        print(f"Errors: {len(errors)}")

        # Save final JSON
        with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
            json.dump(processed_data, f, indent=2, ensure_ascii=False)
        print(f"\n✓ Saved to {OUTPUT_FILE}")

        # Save error log
        if errors:
            with open(ERROR_LOG_FILE, 'w', encoding='utf-8') as f:
                json.dump(errors, f, indent=2, ensure_ascii=False)
            print(f"✓ Error log saved to {ERROR_LOG_FILE}")

        # Generate summary statistics
        print("\n" + "-" * 50)
        print("EXTRACTION SUMMARY:")
        print("-" * 50)

        # Count statistics
        total_with_bio = sum(1 for p in processed_data if p.get("Details", {}).get("ProfileDescription", ""))
        total_with_expertise = sum(1 for p in processed_data if p.get("Expertise", []))
        total_expertise_entries = sum(len(p.get("Expertise", [])) for p in processed_data)
        total_with_facilities = sum(1 for p in processed_data if p.get("Facility_Affiliation", []))
        total_with_web = sum(1 for p in processed_data if p.get("Web_Presence", []))
        total_with_activity = sum(1 for p in processed_data if p.get("OCIP_Activity", []))

        print(f"   Experts with Bio: {total_with_bio}")
        print(f"   Experts with Expertise: {total_with_expertise} ({total_expertise_entries} total entries)")
        print(f"   Experts with Facility Affiliation: {total_with_facilities}")
        print(f"   Experts with Web Presence: {total_with_web}")
        print(f"   Experts with OCIP Activity: {total_with_activity}")

        # Clean up checkpoint file on success
        if os.path.exists(CHECKPOINT_FILE) and len(errors) == 0:
            os.remove(CHECKPOINT_FILE)
            print("\n✓ Checkpoint file cleaned up")

    except KeyboardInterrupt:
        print("\n\n⚠ Script interrupted by user")
        save_checkpoint(processed_data, i + 1 if 'i' in dir() else start_index, errors)
        print(f"   Progress saved. Processed {len(processed_data)} experts so far.")

    except Exception as e:
        print(f"\n✗ CRITICAL ERROR: {e}")
        import traceback
        traceback.print_exc()

        # Emergency save
        if processed_data:
            emergency_file = f"emergency_phase2_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
            with open(emergency_file, 'w', encoding='utf-8') as f:
                json.dump(processed_data, f, indent=2)
            print(f"   Emergency backup saved to {emergency_file}")

    finally:
        print("\n" + "=" * 60)
        print("Script finished. Browser left open for inspection.")
        print("=" * 60)


if __name__ == "__main__":
    main()