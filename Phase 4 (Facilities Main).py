"""
PHASE 4: Facility Deep Scraper
==============================
This script visits each facility's detail page (from Phase 3 master list)
and extracts all information from every accordion section.

Author: AI Assistant
Version: 1.0
Based on Phase 2 Framework
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
INPUT_FILE = "facilities_master_list.json"
OUTPUT_FILE = "facilities_full_details.json"
CHECKPOINT_FILE = "phase4_checkpoint.json"
ERROR_LOG_FILE = "phase4_errors.json"

# Timing Configuration
PAGE_LOAD_WAIT = 2.0
ACCORDION_EXPAND_WAIT = 0.5
BETWEEN_FACILITIES_DELAY = 1.0
REQUEST_DELAY = 0.3

# Rate limiting
BATCH_SIZE = 50
BATCH_PAUSE = 10


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
    """Load the master list from Phase 3."""
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            data = json.load(f)
        print(f"✓ Loaded {len(data)} facilities from {filepath}")
        return data
    except FileNotFoundError:
        print(f"✗ ERROR: {filepath} not found. Run Phase 3 first.")
        return None
    except json.JSONDecodeError as e:
        print(f"✗ ERROR: Invalid JSON in {filepath}: {e}")
        return None


def save_checkpoint(processed_data, current_index, errors):
    """Save progress checkpoint."""
    checkpoint = {
        "timestamp": datetime.now().isoformat(),
        "current_index": current_index,
        "facilities_processed": len(processed_data),
        "errors_count": len(errors),
        "data": processed_data
    }
    with open(CHECKPOINT_FILE, 'w', encoding='utf-8') as f:
        json.dump(checkpoint, f, indent=2, ensure_ascii=False)

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
    text = re.sub(r'\s+', ' ', text.strip())
    text = text.replace('\xa0', ' ')
    return text


def parse_yes_no(element):
    """Parse Yes/No icon elements."""
    try:
        icon = element.find_element(By.CSS_SELECTOR, "span.k-icon")
        title = icon.get_attribute("title")
        if title:
            return title
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
            "li.k-panelbar-header[aria-expanded='false'] > a.k-link, " +
            "li.k-panelbar-item[aria-expanded='false'] > a.k-link"
        )

        expanded_count = 0
        for header in collapsed_headers:
            try:
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", header)
                time.sleep(0.1)

                try:
                    header.click()
                except ElementClickInterceptedException:
                    driver.execute_script("arguments[0].click();", header)

                expanded_count += 1
                time.sleep(ACCORDION_EXPAND_WAIT)

            except StaleElementReferenceException:
                continue
            except Exception:
                continue

        time.sleep(0.5)
        return True

    except Exception as e:
        print(f"      [Warning] Accordion expansion error: {e}")
        return False


def extract_table_grid_data(panel, grid_id=None):
    """
    Generic function to extract data from a Kendo grid table within a panel.
    Returns a list of dictionaries.
    """
    data_list = []

    try:
        # Find the grid
        if grid_id:
            grid = panel.find_element(By.ID, grid_id)
        else:
            grid = panel.find_element(By.CSS_SELECTOR, "div.k-grid")

        # Check for "No records"
        try:
            no_data = grid.find_element(By.CSS_SELECTOR, "tr.k-no-data, div.k-grid-norecords-template")
            if no_data and no_data.is_displayed():
                return []
        except NoSuchElementException:
            pass

        # Get headers
        headers = []
        try:
            header_cells = grid.find_elements(By.CSS_SELECTOR, "thead th")
            for th in header_cells:
                header_text = clean_text(th.text)
                if header_text:
                    headers.append(header_text.replace(" ", "_"))
        except:
            pass

        # Get rows
        rows = grid.find_elements(By.CSS_SELECTOR, "tbody tr.k-master-row")

        for row in rows:
            try:
                cells = row.find_elements(By.TAG_NAME, "td")
                row_data = {}

                for idx, cell in enumerate(cells):
                    # Use header if available, otherwise use index
                    key = headers[idx] if idx < len(headers) else f"Column_{idx}"

                    # Check for links
                    try:
                        link = cell.find_element(By.TAG_NAME, "a")
                        row_data[key] = clean_text(link.text)
                        row_data[f"{key}_URL"] = link.get_attribute("href")
                    except NoSuchElementException:
                        # Check for Yes/No icons
                        try:
                            icon = cell.find_element(By.CSS_SELECTOR, "span.k-icon")
                            row_data[key] = parse_yes_no(cell)
                        except NoSuchElementException:
                            row_data[key] = clean_text(cell.text)

                if row_data:
                    data_list.append(row_data)

            except Exception:
                continue

    except NoSuchElementException:
        pass
    except Exception as e:
        pass

    return data_list


def extract_key_value_pairs(panel):
    """
    Generic function to extract label-value pairs from a panel.
    Returns a dictionary.
    """
    data = {}

    try:
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
                value_cols = row.find_elements(By.CSS_SELECTOR,
                    "div.col-md-9, div.col-md-7, div.col-md-8, div.col-md-10")

                if not value_cols:
                    # Try finding any column that's not the label column
                    all_cols = row.find_elements(By.CSS_SELECTOR, "div[class*='col-']")
                    for col in all_cols:
                        if label not in col.find_elements(By.TAG_NAME, "label"):
                            value_cols = [col]
                            break

                if not value_cols:
                    continue

                value_col = value_cols[0]
                field_key = label_for if label_for else label_text.replace(" ", "_").replace(":", "")

                # Handle different field types
                # Check for Yes/No icons
                try:
                    icon = value_col.find_element(By.CSS_SELECTOR, "span.k-icon")
                    data[field_key] = parse_yes_no(value_col)
                    continue
                except NoSuchElementException:
                    pass

                # Check for links (email, phone, URL)
                try:
                    link = value_col.find_element(By.TAG_NAME, "a")
                    href = link.get_attribute("href") or ""
                    if href.startswith("mailto:"):
                        data["Email"] = clean_text(link.text)
                    elif href.startswith("tel:"):
                        data["Phone"] = clean_text(link.text)
                    else:
                        data[field_key] = clean_text(link.text)
                        data[f"{field_key}_URL"] = href
                    continue
                except NoSuchElementException:
                    pass

                # Check for rating
                try:
                    rating = value_col.find_element(By.CSS_SELECTOR, "span.k-rating")
                    rating_value = rating.get_attribute("aria-valuenow")
                    data[field_key] = rating_value if rating_value else "Not Rated"
                    continue
                except NoSuchElementException:
                    pass

                # Default: plain text
                data[field_key] = clean_text(value_col.text)

            except Exception:
                continue

    except Exception as e:
        pass

    return data


# ==========================================
# SECTION EXTRACTORS (12 Sections)
# ==========================================

def extract_general_information(driver):
    """Extract data from General Information section (li[1])."""
    data = {}

    try:
        # Find panel by index - first panel
        panels = driver.find_elements(By.CSS_SELECTOR, "ul.k-panelbar > li")
        if len(panels) < 1:
            return data

        panel = panels[0]

        # Extract key-value pairs
        data = extract_key_value_pairs(panel)

        # Extract breadcrumb if present (Academic Unit path)
        try:
            breadcrumb_items = panel.find_elements(By.CSS_SELECTOR, "ol.breadcrumb li")
            if breadcrumb_items:
                path = [clean_text(item.text) for item in breadcrumb_items if clean_text(item.text)]
                data["Academic_Unit_Path"] = " > ".join(path)
        except:
            pass

        # Extract image URL if present
        try:
            img = panel.find_element(By.CSS_SELECTOR, "img[alt]")
            data["Image_URL"] = img.get_attribute("src")
        except:
            pass

    except Exception as e:
        print(f"      [Warning] General Information extraction error: {e}")

    return data


def extract_academic_unit_details(driver):
    """Extract data from Academic Unit Details section (li[2])."""
    data = {}

    try:
        panels = driver.find_elements(By.CSS_SELECTOR, "ul.k-panelbar > li")
        if len(panels) < 2:
            return data

        panel = panels[1]
        data = extract_key_value_pairs(panel)

    except Exception as e:
        print(f"      [Warning] Academic Unit Details extraction error: {e}")

    return data


def extract_provinces_served(driver):
    """Extract data from Provinces Served section (li[3])."""
    provinces_list = []

    try:
        panels = driver.find_elements(By.CSS_SELECTOR, "ul.k-panelbar > li")
        if len(panels) < 3:
            return provinces_list

        panel = panels[2]

        # Could be a table/grid or a list
        grid_data = extract_table_grid_data(panel)
        if grid_data:
            return grid_data

        # Try extracting as list items
        try:
            items = panel.find_elements(By.CSS_SELECTOR, "li, div.item, span.tag")
            for item in items:
                text = clean_text(item.text)
                if text:
                    provinces_list.append(text)
        except:
            pass

        # Try extracting as plain text
        if not provinces_list:
            try:
                content_div = panel.find_element(By.CSS_SELECTOR, "div.k-content, div.panel-body")
                text = clean_text(content_div.text)
                if text:
                    provinces_list = [p.strip() for p in text.split(',') if p.strip()]
            except:
                pass

    except Exception as e:
        print(f"      [Warning] Provinces Served extraction error: {e}")

    return provinces_list


def extract_activities_offered(driver):
    """Extract data from Activities Offered section (li[4])."""
    activities_list = []

    try:
        panels = driver.find_elements(By.CSS_SELECTOR, "ul.k-panelbar > li")
        if len(panels) < 4:
            return activities_list

        panel = panels[3]

        # Try grid extraction first
        grid_data = extract_table_grid_data(panel)
        if grid_data:
            return grid_data

        # Try list extraction
        try:
            items = panel.find_elements(By.CSS_SELECTOR, "li, div.item, span.tag, div.chip")
            for item in items:
                text = clean_text(item.text)
                if text and text not in ["Activities Offered", ""]:
                    activities_list.append(text)
        except:
            pass

    except Exception as e:
        print(f"      [Warning] Activities Offered extraction error: {e}")

    return activities_list


def extract_sectors_served(driver):
    """Extract data from Sectors Served section (li[5])."""
    sectors_list = []

    try:
        panels = driver.find_elements(By.CSS_SELECTOR, "ul.k-panelbar > li")
        if len(panels) < 5:
            return sectors_list

        panel = panels[4]

        # Try grid extraction
        grid_data = extract_table_grid_data(panel)
        if grid_data:
            return grid_data

        # Try list/tag extraction
        try:
            items = panel.find_elements(By.CSS_SELECTOR, "li, div.item, span.tag, div.chip")
            for item in items:
                text = clean_text(item.text)
                if text and text not in ["Sectors Served", ""]:
                    sectors_list.append(text)
        except:
            pass

    except Exception as e:
        print(f"      [Warning] Sectors Served extraction error: {e}")

    return sectors_list


def extract_contacts(driver):
    """Extract data from Contacts section (li[6])."""
    contacts_list = []

    try:
        panels = driver.find_elements(By.CSS_SELECTOR, "ul.k-panelbar > li")
        if len(panels) < 6:
            return contacts_list

        panel = panels[5]

        # Try grid extraction (contacts usually in a table)
        grid_data = extract_table_grid_data(panel)
        if grid_data:
            return grid_data

        # Fallback: Extract contact cards
        try:
            contact_cards = panel.find_elements(By.CSS_SELECTOR, "div.contact-card, div.card, div.row")
            for card in contact_cards:
                contact = {}

                # Name
                try:
                    name_elem = card.find_element(By.CSS_SELECTOR, "h4, h5, .name, strong")
                    contact["Name"] = clean_text(name_elem.text)
                except:
                    pass

                # Email
                try:
                    email_elem = card.find_element(By.CSS_SELECTOR, "a[href^='mailto:']")
                    contact["Email"] = clean_text(email_elem.text)
                except:
                    pass

                # Phone
                try:
                    phone_elem = card.find_element(By.CSS_SELECTOR, "a[href^='tel:']")
                    contact["Phone"] = clean_text(phone_elem.text)
                except:
                    pass

                # Role/Title
                try:
                    role_elem = card.find_element(By.CSS_SELECTOR, ".role, .title, .position")
                    contact["Role"] = clean_text(role_elem.text)
                except:
                    pass

                if contact:
                    contacts_list.append(contact)

        except:
            pass

    except Exception as e:
        print(f"      [Warning] Contacts extraction error: {e}")

    return contacts_list


def extract_locations(driver):
    """Extract data from Locations section (li[7])."""
    locations_list = []

    try:
        panels = driver.find_elements(By.CSS_SELECTOR, "ul.k-panelbar > li")
        if len(panels) < 7:
            return locations_list

        panel = panels[6]

        # Try grid extraction
        grid_data = extract_table_grid_data(panel)
        if grid_data:
            return grid_data

        # Fallback: Extract address blocks
        try:
            address_blocks = panel.find_elements(By.CSS_SELECTOR, "address, div.address, div.location")
            for block in address_blocks:
                location = {
                    "Address": clean_text(block.text)
                }
                locations_list.append(location)
        except:
            pass

        # Another fallback: key-value pairs
        if not locations_list:
            kv_data = extract_key_value_pairs(panel)
            if kv_data:
                locations_list.append(kv_data)

    except Exception as e:
        print(f"      [Warning] Locations extraction error: {e}")

    return locations_list


def extract_facility_descriptors(driver):
    """Extract data from Facility Descriptors section (li[8])."""
    descriptors = {}

    try:
        panels = driver.find_elements(By.CSS_SELECTOR, "ul.k-panelbar > li")
        if len(panels) < 8:
            return descriptors

        panel = panels[7]

        # Try key-value extraction
        descriptors = extract_key_value_pairs(panel)

        # Also try grid extraction for list-like descriptors
        grid_data = extract_table_grid_data(panel)
        if grid_data:
            descriptors["Descriptors_List"] = grid_data

        # Try extracting tags/chips
        try:
            tags = panel.find_elements(By.CSS_SELECTOR, "span.tag, div.chip, span.badge")
            if tags:
                descriptors["Tags"] = [clean_text(tag.text) for tag in tags if clean_text(tag.text)]
        except:
            pass

    except Exception as e:
        print(f"      [Warning] Facility Descriptors extraction error: {e}")

    return descriptors


def extract_languages_serviced(driver):
    """Extract data from Languages Serviced section (li[9])."""
    languages_list = []

    try:
        panels = driver.find_elements(By.CSS_SELECTOR, "ul.k-panelbar > li")
        if len(panels) < 9:
            return languages_list

        panel = panels[8]

        # Try grid extraction
        grid_data = extract_table_grid_data(panel)
        if grid_data:
            return grid_data

        # Try list/tag extraction
        try:
            items = panel.find_elements(By.CSS_SELECTOR, "li, span.tag, div.chip, span.badge")
            for item in items:
                text = clean_text(item.text)
                if text and text not in ["Languages Serviced", ""]:
                    languages_list.append(text)
        except:
            pass

        # Fallback: plain text
        if not languages_list:
            try:
                content = panel.find_element(By.CSS_SELECTOR, "div.k-content, div.panel-body")
                text = clean_text(content.text)
                if text:
                    languages_list = [l.strip() for l in text.split(',') if l.strip()]
            except:
                pass

    except Exception as e:
        print(f"      [Warning] Languages Serviced extraction error: {e}")

    return languages_list


def extract_web_presence(driver):
    """Extract data from Web Presence section (li[10])."""
    web_presence_list = []

    try:
        panels = driver.find_elements(By.CSS_SELECTOR, "ul.k-panelbar > li")
        if len(panels) < 10:
            return web_presence_list

        panel = panels[9]

        # Try grid extraction
        grid_data = extract_table_grid_data(panel)
        if grid_data:
            return grid_data

        # Fallback: find all links
        try:
            links = panel.find_elements(By.TAG_NAME, "a")
            for link in links:
                href = link.get_attribute("href")
                text = clean_text(link.text)
                if href and not href.startswith("javascript"):
                    web_presence_list.append({
                        "Name": text if text else "Link",
                        "URL": href
                    })
        except:
            pass

    except Exception as e:
        print(f"      [Warning] Web Presence extraction error: {e}")

    return web_presence_list


def extract_ocip_activity(driver):
    """Extract data from OCIP Activity section (li[11])."""
    activity_list = []

    try:
        panels = driver.find_elements(By.CSS_SELECTOR, "ul.k-panelbar > li")
        if len(panels) < 11:
            return activity_list

        panel = panels[10]

        # Try grid extraction
        grid_data = extract_table_grid_data(panel)
        if grid_data:
            return grid_data

        # Fallback: key-value pairs
        kv_data = extract_key_value_pairs(panel)
        if kv_data:
            activity_list.append(kv_data)

    except Exception as e:
        print(f"      [Warning] OCIP Activity extraction error: {e}")

    return activity_list


def extract_audit_trail(driver):
    """Extract data from Audit Trail section (li[12])."""
    data = {}

    try:
        panels = driver.find_elements(By.CSS_SELECTOR, "ul.k-panelbar > li")
        if len(panels) < 12:
            return data

        panel = panels[11]

        # Extract key-value pairs
        data = extract_key_value_pairs(panel)

    except Exception as e:
        print(f"      [Warning] Audit Trail extraction error: {e}")

    return data


# ==========================================
# MAIN EXTRACTION FUNCTION
# ==========================================

def extract_facility_full_profile(driver, wait, facility_basic_info):
    """
    Extract all information from a facility's detail page.
    Returns a complete profile dictionary.
    """
    url = facility_basic_info.get("Manage_URL", "")

    if not url or url == "Not Found":
        return None

    try:
        # Navigate to the detail page
        driver.get(url)
        time.sleep(PAGE_LOAD_WAIT)

        # Wait for page to load
        try:
            wait.until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, "ul.k-panelbar, li.k-panelbar-header, li.k-panelbar-item")
            ))
        except TimeoutException:
            print(f"      [Warning] Page load timeout for {url}")
            return None

        # Expand all accordion sections
        expand_all_accordions(driver, wait)

        # Initialize profile with basic info from Phase 3
        profile = {
            "Meta": {
                "Source_URL": url,
                "Scraped_At": datetime.now().isoformat(),
                "Institution": facility_basic_info.get("Institution", ""),
                "Facility_Name_From_List": facility_basic_info.get("Facility_Name", ""),
                "Facility_ID": facility_basic_info.get("Facility_ID", ""),
                "Type_From_List": facility_basic_info.get("Type", "")
            }
        }

        # Extract each section (12 sections total)
        print("         Extracting sections...")

        profile["General_Information"] = extract_general_information(driver)
        time.sleep(REQUEST_DELAY)

        profile["Academic_Unit_Details"] = extract_academic_unit_details(driver)
        time.sleep(REQUEST_DELAY)

        profile["Provinces_Served"] = extract_provinces_served(driver)
        time.sleep(REQUEST_DELAY)

        profile["Activities_Offered"] = extract_activities_offered(driver)
        time.sleep(REQUEST_DELAY)

        profile["Sectors_Served"] = extract_sectors_served(driver)
        time.sleep(REQUEST_DELAY)

        profile["Contacts"] = extract_contacts(driver)
        time.sleep(REQUEST_DELAY)

        profile["Locations"] = extract_locations(driver)
        time.sleep(REQUEST_DELAY)

        profile["Facility_Descriptors"] = extract_facility_descriptors(driver)
        time.sleep(REQUEST_DELAY)

        profile["Languages_Serviced"] = extract_languages_serviced(driver)
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
    """Main execution function for Phase 4."""

    print("\n" + "=" * 60)
    print("  PHASE 4: FACILITY DEEP SCRAPER")
    print("  Extracting Full Profiles from Detail Pages")
    print("=" * 60)

    # Load master list from Phase 3
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
        print(f"   Progress: {checkpoint['current_index']}/{len(master_list)} facilities")
        print(f"   Successfully processed: {checkpoint['facilities_processed']}")
        print(f"   Errors: {checkpoint['errors_count']}")

        resume = input("\nResume from checkpoint? (y/n): ").strip().lower()
        if resume == 'y':
            processed_data = checkpoint['data']
            start_index = checkpoint['current_index']
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

        # Process each facility
        print("\n" + "-" * 50)
        print("STEP 2: EXTRACTING FACILITY PROFILES")
        print("-" * 50)

        total = len(master_list)

        for i, facility in enumerate(master_list[start_index:], start=start_index):
            facility_name = facility.get("Facility_Name", "Unknown")
            institution = facility.get("Institution", "Unknown")
            url = facility.get("Manage_URL", "")

            print(f"\n[{i + 1}/{total}] {facility_name} ({institution})")

            if not url or url == "Not Found":
                print("      → Skipped: No valid URL")
                errors.append({
                    "index": i,
                    "name": facility_name,
                    "institution": institution,
                    "reason": "No valid Manage URL"
                })
                continue

            # Extract full profile
            profile = extract_facility_full_profile(driver, wait, facility)

            if profile:
                processed_data.append(profile)
                print(f"      ✓ Extracted successfully")

                # Show summary of what was found
                provinces_count = len(profile.get("Provinces_Served", []))
                activities_count = len(profile.get("Activities_Offered", []))
                contacts_count = len(profile.get("Contacts", []))
                locations_count = len(profile.get("Locations", []))

                print(f"         Provinces: {provinces_count} | Activities: {activities_count} | "
                      f"Contacts: {contacts_count} | Locations: {locations_count}")
            else:
                errors.append({
                    "index": i,
                    "name": facility_name,
                    "institution": institution,
                    "url": url,
                    "reason": "Extraction failed"
                })
                print(f"      ✗ Extraction failed")

            # Save checkpoint periodically
            if (i + 1) % 10 == 0:
                save_checkpoint(processed_data, i + 1, errors)
                print(f"\n   [Checkpoint saved: {len(processed_data)} profiles]")

            # Rate limiting
            if (i + 1) % BATCH_SIZE == 0 and i + 1 < total:
                print(f"\n   [Rate limit pause: {BATCH_PAUSE}s...]")
                time.sleep(BATCH_PAUSE)

            # Delay between facilities
            time.sleep(BETWEEN_FACILITIES_DELAY)

        # ===== FINAL SAVE =====
        print("\n" + "=" * 60)
        print("SCRAPING COMPLETE!")
        print("=" * 60)
        print(f"Total Facilities Processed: {len(processed_data)}")
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

        total_with_provinces = sum(1 for p in processed_data if p.get("Provinces_Served", []))
        total_with_activities = sum(1 for p in processed_data if p.get("Activities_Offered", []))
        total_with_contacts = sum(1 for p in processed_data if p.get("Contacts", []))
        total_with_locations = sum(1 for p in processed_data if p.get("Locations", []))
        total_with_sectors = sum(1 for p in processed_data if p.get("Sectors_Served", []))
        total_with_languages = sum(1 for p in processed_data if p.get("Languages_Serviced", []))
        total_with_web = sum(1 for p in processed_data if p.get("Web_Presence", []))
        total_with_ocip = sum(1 for p in processed_data if p.get("OCIP_Activity", []))

        print(f"   Facilities with Provinces Served: {total_with_provinces}")
        print(f"   Facilities with Activities Offered: {total_with_activities}")
        print(f"   Facilities with Contacts: {total_with_contacts}")
        print(f"   Facilities with Locations: {total_with_locations}")
        print(f"   Facilities with Sectors Served: {total_with_sectors}")
        print(f"   Facilities with Languages: {total_with_languages}")
        print(f"   Facilities with Web Presence: {total_with_web}")
        print(f"   Facilities with OCIP Activity: {total_with_ocip}")

        # Clean up checkpoint file on success
        if os.path.exists(CHECKPOINT_FILE) and len(errors) == 0:
            os.remove(CHECKPOINT_FILE)
            print("\n✓ Checkpoint file cleaned up")

    except KeyboardInterrupt:
        print("\n\n⚠ Script interrupted by user")
        save_checkpoint(processed_data, i + 1 if 'i' in dir() else start_index, errors)
        print(f"   Progress saved. Processed {len(processed_data)} facilities so far.")

    except Exception as e:
        print(f"\n✗ CRITICAL ERROR: {e}")
        import traceback
        traceback.print_exc()

        # Emergency save
        if processed_data:
            emergency_file = f"emergency_phase4_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
            with open(emergency_file, 'w', encoding='utf-8') as f:
                json.dump(processed_data, f, indent=2)
            print(f"   Emergency backup saved to {emergency_file}")

    finally:
        print("\n" + "=" * 60)
        print("Script finished. Browser left open for inspection.")
        print("=" * 60)


if __name__ == "__main__":
    main()