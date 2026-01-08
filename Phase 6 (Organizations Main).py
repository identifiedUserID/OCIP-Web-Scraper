"""
PHASE 6: Organization Deep Scraper
==================================
This script visits each organization's detail page (from Phase 5 master list)
and extracts all information from every accordion section.

* Checkpoints are saved after EVERY organization to ensure real-time data access. *

Author: AI Assistant
Version: 1.0
Based on Phase 4 Framework
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
INPUT_FILE = "organizations_master_list.json"
OUTPUT_FILE = "organizations_full_details.json"
CHECKPOINT_FILE = "phase6_checkpoint.json"
ERROR_LOG_FILE = "phase6_errors.json"

# Timing Configuration
PAGE_LOAD_WAIT = 2.0
ACCORDION_EXPAND_WAIT = 0.5
BETWEEN_ORGS_DELAY = 1.0
REQUEST_DELAY = 0.3

# Rate limiting
BATCH_SIZE = 50
BATCH_PAUSE = 10


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
def load_master_list(filepath):
    """Load the master list from Phase 5."""
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            data = json.load(f)
        print(f"✓ Loaded {len(data)} organizations from {filepath}")
        return data
    except FileNotFoundError:
        print(f"✗ ERROR: {filepath} not found. Run Phase 5 first.")
        return None
    except json.JSONDecodeError as e:
        print(f"✗ ERROR: Invalid JSON in {filepath}: {e}")
        return None


def save_checkpoint(processed_data, current_index, errors, total):
    """
    Save progress checkpoint IMMEDIATELY to disk.
    File is accessible in real-time during script execution.
    """
    checkpoint = {
        "timestamp": datetime.now().isoformat(),
        "current_index": current_index,
        "total_organizations": total,
        "organizations_processed": len(processed_data),
        "errors_count": len(errors),
        "progress_percent": round((current_index / total) * 100, 2) if total > 0 else 0,
        "data": processed_data
    }
    try:
        with open(CHECKPOINT_FILE, 'w', encoding='utf-8') as f:
            json.dump(checkpoint, f, indent=2, ensure_ascii=False)
            f.flush()
            os.fsync(f.fileno())  # Force write to disk
    except Exception as e:
        print(f"      [Warning] Checkpoint save failed: {e}")

    # Also save errors immediately
    try:
        with open(ERROR_LOG_FILE, 'w', encoding='utf-8') as f:
            json.dump(errors, f, indent=2, ensure_ascii=False)
            f.flush()
            os.fsync(f.fileno())
    except:
        pass


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
        if "k-i-checkbox-checked" in classes or "k-i-check" in classes:
            return "Yes"
        elif "k-i-checkbox" in classes or "k-i-close" in classes:
            return "No"
    except:
        pass
    return clean_text(element.text)


def expand_all_accordions(driver, wait):
    """Expand all collapsed accordion panels."""
    try:
        collapsed_headers = driver.find_elements(
            By.CSS_SELECTOR,
            "li.k-panelbar-header[aria-expanded='false'] > a.k-link, " +
            "li.k-panelbar-item[aria-expanded='false'] > a.k-link, " +
            "li[aria-expanded='false'] > a.k-link"
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
        print(f"         Expanded {expanded_count} accordion sections")
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
            try:
                grid = panel.find_element(By.ID, grid_id)
            except NoSuchElementException:
                grid = panel.find_element(By.CSS_SELECTOR, "div.k-grid")
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
                    headers.append(header_text.replace(" ", "_").replace("?", ""))
        except:
            pass

        # Get rows
        rows = grid.find_elements(By.CSS_SELECTOR, "tbody tr.k-master-row, tbody tr:not(.k-detail-row)")

        for row in rows:
            try:
                cells = row.find_elements(By.TAG_NAME, "td")
                row_data = {}

                for idx, cell in enumerate(cells):
                    key = headers[idx] if idx < len(headers) else f"Column_{idx}"

                    # Skip empty keys
                    if not key or key == "":
                        key = f"Column_{idx}"

                    # Check for links
                    try:
                        link = cell.find_element(By.TAG_NAME, "a")
                        href = link.get_attribute("href") or ""
                        link_text = clean_text(link.text)

                        if href.startswith("mailto:"):
                            row_data["Email"] = link_text
                        elif href.startswith("tel:"):
                            row_data["Phone"] = link_text
                        else:
                            row_data[key] = link_text
                            if href and not href.startswith("javascript"):
                                row_data[f"{key}_URL"] = href
                        continue
                    except NoSuchElementException:
                        pass

                    # Check for Yes/No icons
                    try:
                        icon = cell.find_element(By.CSS_SELECTOR, "span.k-icon")
                        row_data[key] = parse_yes_no(cell)
                        continue
                    except NoSuchElementException:
                        pass

                    # Default: plain text
                    row_data[key] = clean_text(cell.text)

                if row_data and any(v for v in row_data.values() if v):
                    data_list.append(row_data)

            except Exception:
                continue

    except NoSuchElementException:
        pass
    except Exception:
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
                label_text = clean_text(label.text).replace(":", "")

                # Find value column
                value_cols = row.find_elements(By.CSS_SELECTOR,
                    "div.col-md-9, div.col-md-7, div.col-md-8, div.col-md-10, div.col-md-6")

                if not value_cols:
                    all_cols = row.find_elements(By.CSS_SELECTOR, "div[class*='col-']")
                    for col in all_cols:
                        col_labels = col.find_elements(By.TAG_NAME, "label")
                        if not col_labels or label not in col_labels:
                            if col.text.strip():
                                value_cols = [col]
                                break

                if not value_cols:
                    continue

                value_col = value_cols[0]
                field_key = label_for if label_for else label_text.replace(" ", "_")

                # Skip empty keys
                if not field_key:
                    continue

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
                    links = value_col.find_elements(By.TAG_NAME, "a")
                    if links:
                        for link in links:
                            href = link.get_attribute("href") or ""
                            link_text = clean_text(link.text)
                            if href.startswith("mailto:"):
                                data["Email"] = link_text
                            elif href.startswith("tel:"):
                                data["Phone"] = link_text
                            else:
                                data[field_key] = link_text
                                if href and not href.startswith("javascript"):
                                    data[f"{field_key}_URL"] = href
                        continue
                except:
                    pass

                # Check for rating
                try:
                    rating = value_col.find_element(By.CSS_SELECTOR, "span.k-rating")
                    rating_value = rating.get_attribute("aria-valuenow")
                    data[field_key] = rating_value if rating_value else "Not Rated"
                    continue
                except NoSuchElementException:
                    pass

                # Check for breadcrumb
                try:
                    breadcrumb = value_col.find_element(By.CSS_SELECTOR, "ol.breadcrumb")
                    items = breadcrumb.find_elements(By.TAG_NAME, "li")
                    path = [clean_text(item.text) for item in items if clean_text(item.text)]
                    data[field_key] = " > ".join(path)
                    continue
                except NoSuchElementException:
                    pass

                # Default: plain text
                text_value = clean_text(value_col.text)
                if text_value:
                    data[field_key] = text_value

            except Exception:
                continue

    except Exception:
        pass

    return data


# ==========================================
# SECTION EXTRACTORS (10 Sections)
# ==========================================

def extract_general_information(driver):
    """Extract data from General Information section (li[1])."""
    data = {}

    try:
        panels = driver.find_elements(By.CSS_SELECTOR, "ul.k-panelbar > li")
        if len(panels) < 1:
            return data

        panel = panels[0]
        data = extract_key_value_pairs(panel)

        # Extract image URL if present
        try:
            img = panel.find_element(By.CSS_SELECTOR, "img[alt], img.logo, img.org-image")
            src = img.get_attribute("src")
            if src:
                data["Image_URL"] = src
        except:
            pass

        # Extract any additional standalone text blocks
        try:
            description_divs = panel.find_elements(By.CSS_SELECTOR, "div.description, div.summary, p.lead")
            for div in description_divs:
                text = clean_text(div.text)
                if text and len(text) > 50:
                    data["Description"] = text
                    break
        except:
            pass

    except Exception as e:
        print(f"      [Warning] General Information extraction error: {e}")

    return data


def extract_organization_information(driver):
    """Extract data from Organization Information section (li[2])."""
    data = {}

    try:
        panels = driver.find_elements(By.CSS_SELECTOR, "ul.k-panelbar > li")
        if len(panels) < 2:
            return data

        panel = panels[1]
        data = extract_key_value_pairs(panel)

        # Also check for any grids in this section
        grid_data = extract_table_grid_data(panel)
        if grid_data:
            data["Additional_Info"] = grid_data

    except Exception as e:
        print(f"      [Warning] Organization Information extraction error: {e}")

    return data


def extract_annual_information(driver):
    """Extract data from Annual Information section (li[3])."""
    data = {}

    try:
        panels = driver.find_elements(By.CSS_SELECTOR, "ul.k-panelbar > li")
        if len(panels) < 3:
            return data

        panel = panels[2]

        # This section likely contains financial/annual data
        # Try key-value first
        data = extract_key_value_pairs(panel)

        # Also try grid extraction for tabular annual data
        grid_data = extract_table_grid_data(panel)
        if grid_data:
            data["Annual_Records"] = grid_data

    except Exception as e:
        print(f"      [Warning] Annual Information extraction error: {e}")

    return data


def extract_naics_sectors(driver):
    """Extract data from NAICS Sectors section (li[4])."""
    sectors_list = []

    try:
        panels = driver.find_elements(By.CSS_SELECTOR, "ul.k-panelbar > li")
        if len(panels) < 4:
            return sectors_list

        panel = panels[3]

        # Try grid extraction first (NAICS codes often in table)
        grid_data = extract_table_grid_data(panel)
        if grid_data:
            return grid_data

        # Try list extraction
        try:
            items = panel.find_elements(By.CSS_SELECTOR, "li, div.item, span.tag, div.chip, span.badge")
            for item in items:
                text = clean_text(item.text)
                if text and text not in ["NAICS Sectors", ""] and len(text) > 1:
                    sectors_list.append(text)
        except:
            pass

        # Fallback: plain text parsing
        if not sectors_list:
            try:
                content = panel.find_element(By.CSS_SELECTOR, "div.k-content, div.panel-body, div.content")
                text = clean_text(content.text)
                if text:
                    # Split by common delimiters
                    for delimiter in [',', ';', '\n']:
                        if delimiter in text:
                            sectors_list = [s.strip() for s in text.split(delimiter) if s.strip()]
                            break
            except:
                pass

    except Exception as e:
        print(f"      [Warning] NAICS Sectors extraction error: {e}")

    return sectors_list


def extract_contacts(driver):
    """Extract data from Contacts section (li[5])."""
    contacts_list = []

    try:
        panels = driver.find_elements(By.CSS_SELECTOR, "ul.k-panelbar > li")
        if len(panels) < 5:
            return contacts_list

        panel = panels[4]

        # Try grid extraction (contacts usually in a table)
        grid_data = extract_table_grid_data(panel)
        if grid_data:
            return grid_data

        # Fallback: Extract contact cards/blocks
        try:
            contact_blocks = panel.find_elements(By.CSS_SELECTOR,
                "div.contact-card, div.card, div.contact, div.row")

            for block in contact_blocks:
                contact = {}

                # Name
                try:
                    name_elem = block.find_element(By.CSS_SELECTOR, "h4, h5, .name, strong, b")
                    name = clean_text(name_elem.text)
                    if name:
                        contact["Name"] = name
                except:
                    pass

                # Email
                try:
                    email_elem = block.find_element(By.CSS_SELECTOR, "a[href^='mailto:']")
                    contact["Email"] = clean_text(email_elem.text)
                except:
                    pass

                # Phone
                try:
                    phone_elem = block.find_element(By.CSS_SELECTOR, "a[href^='tel:']")
                    contact["Phone"] = clean_text(phone_elem.text)
                except:
                    pass

                # Role/Title
                try:
                    role_elem = block.find_element(By.CSS_SELECTOR, ".role, .title, .position, small")
                    role = clean_text(role_elem.text)
                    if role:
                        contact["Role"] = role
                except:
                    pass

                if contact and len(contact) > 0:
                    contacts_list.append(contact)

        except:
            pass

    except Exception as e:
        print(f"      [Warning] Contacts extraction error: {e}")

    return contacts_list


def extract_locations(driver):
    """Extract data from Locations section (li[6])."""
    locations_list = []

    try:
        panels = driver.find_elements(By.CSS_SELECTOR, "ul.k-panelbar > li")
        if len(panels) < 6:
            return locations_list

        panel = panels[5]

        # Try grid extraction
        grid_data = extract_table_grid_data(panel)
        if grid_data:
            return grid_data

        # Fallback: Extract address blocks
        try:
            address_blocks = panel.find_elements(By.CSS_SELECTOR,
                "address, div.address, div.location, div.row")
            for block in address_blocks:
                text = clean_text(block.text)
                if text and len(text) > 5:
                    location = {"Address": text}

                    # Try to find specific fields
                    try:
                        city = block.find_element(By.CSS_SELECTOR, ".city")
                        location["City"] = clean_text(city.text)
                    except:
                        pass

                    try:
                        province = block.find_element(By.CSS_SELECTOR, ".province, .state")
                        location["Province"] = clean_text(province.text)
                    except:
                        pass

                    try:
                        postal = block.find_element(By.CSS_SELECTOR, ".postal, .zip")
                        location["Postal_Code"] = clean_text(postal.text)
                    except:
                        pass

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


def extract_languages_serviced(driver):
    """Extract data from Languages Serviced section (li[7])."""
    languages_list = []

    try:
        panels = driver.find_elements(By.CSS_SELECTOR, "ul.k-panelbar > li")
        if len(panels) < 7:
            return languages_list

        panel = panels[6]

        # Try grid extraction
        grid_data = extract_table_grid_data(panel)
        if grid_data:
            return grid_data

        # Try list/tag extraction
        try:
            items = panel.find_elements(By.CSS_SELECTOR,
                "li, span.tag, div.chip, span.badge, div.item")
            for item in items:
                text = clean_text(item.text)
                if text and text not in ["Languages Serviced", ""] and len(text) > 1:
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
    """Extract data from Web Presence section (li[8])."""
    web_presence_list = []

    try:
        panels = driver.find_elements(By.CSS_SELECTOR, "ul.k-panelbar > li")
        if len(panels) < 8:
            return web_presence_list

        panel = panels[7]

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

                # Skip navigation/accordion links
                if not href or href.startswith("javascript") or href == "#":
                    continue
                if "k-link" in (link.get_attribute("class") or ""):
                    continue

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
    """Extract data from OCIP Activity section (li[9])."""
    activity_list = []

    try:
        panels = driver.find_elements(By.CSS_SELECTOR, "ul.k-panelbar > li")
        if len(panels) < 9:
            return activity_list

        panel = panels[8]

        # Try grid extraction (OCIP activity usually in table format)
        grid_data = extract_table_grid_data(panel)
        if grid_data:
            return grid_data

        # Fallback: key-value pairs
        kv_data = extract_key_value_pairs(panel)
        if kv_data:
            activity_list.append(kv_data)

        # Try to find project links
        try:
            project_links = panel.find_elements(By.CSS_SELECTOR, "a[href*='Project'], a[href*='Request']")
            for link in project_links:
                href = link.get_attribute("href")
                text = clean_text(link.text)
                if href and text:
                    activity_list.append({
                        "Project_Name": text,
                        "Project_URL": href
                    })
        except:
            pass

    except Exception as e:
        print(f"      [Warning] OCIP Activity extraction error: {e}")

    return activity_list


def extract_audit_trail(driver):
    """Extract data from Audit Trail section (li[10])."""
    data = {}

    try:
        panels = driver.find_elements(By.CSS_SELECTOR, "ul.k-panelbar > li")
        if len(panels) < 10:
            return data

        panel = panels[9]

        # Extract key-value pairs (typically created/modified dates)
        data = extract_key_value_pairs(panel)

        # Also try grid in case there's a history table
        grid_data = extract_table_grid_data(panel)
        if grid_data:
            data["History"] = grid_data

    except Exception as e:
        print(f"      [Warning] Audit Trail extraction error: {e}")

    return data


# ==========================================
# MAIN EXTRACTION FUNCTION
# ==========================================

def extract_organization_full_profile(driver, wait, org_basic_info):
    """
    Extract all information from an organization's detail page.
    Returns a complete profile dictionary.
    """
    url = org_basic_info.get("Manage_URL", "")

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

        # Initialize profile with basic info from Phase 5
        profile = {
            "Meta": {
                "Source_URL": url,
                "Scraped_At": datetime.now().isoformat(),
                "Organization_Name_From_List": org_basic_info.get("Organization_Name", ""),
                "Provinces_From_List": org_basic_info.get("Provinces", ""),
                "Sectors_From_List": org_basic_info.get("Sectors", "")
            }
        }

        # Extract each section (10 sections total)
        print("         Extracting sections...")

        profile["General_Information"] = extract_general_information(driver)
        time.sleep(REQUEST_DELAY)

        profile["Organization_Information"] = extract_organization_information(driver)
        time.sleep(REQUEST_DELAY)

        profile["Annual_Information"] = extract_annual_information(driver)
        time.sleep(REQUEST_DELAY)

        profile["NAICS_Sectors"] = extract_naics_sectors(driver)
        time.sleep(REQUEST_DELAY)

        profile["Contacts"] = extract_contacts(driver)
        time.sleep(REQUEST_DELAY)

        profile["Locations"] = extract_locations(driver)
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
    """Main execution function for Phase 6."""

    print("\n" + "=" * 60)
    print("  PHASE 6: ORGANIZATION DEEP SCRAPER")
    print("  Extracting Full Profiles from Detail Pages")
    print("=" * 60)
    print("\n  *** Checkpoints saved after EVERY organization ***")
    print(f"  *** Open '{CHECKPOINT_FILE}' anytime to see progress ***\n")

    # Load master list from Phase 5
    master_list = load_master_list(INPUT_FILE)
    if not master_list:
        return

    # Initialize
    driver = get_driver()
    wait = WebDriverWait(driver, 20)
    processed_data = []
    errors = []
    start_index = 0
    total = len(master_list)

    # Check for existing checkpoint
    checkpoint = load_checkpoint()
    if checkpoint:
        print(f"\n⚠ Found checkpoint from {checkpoint['timestamp']}")
        print(f"   Progress: {checkpoint['current_index']}/{checkpoint.get('total_organizations', total)} organizations")
        print(f"   Successfully processed: {checkpoint['organizations_processed']}")
        print(f"   Errors: {checkpoint['errors_count']}")
        print(f"   Completion: {checkpoint.get('progress_percent', 0)}%")

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

        # Process each organization
        print("\n" + "-" * 50)
        print("STEP 2: EXTRACTING ORGANIZATION PROFILES")
        print("-" * 50)

        for i, org in enumerate(master_list[start_index:], start=start_index):
            org_name = org.get("Organization_Name", "Unknown")
            url = org.get("Manage_URL", "")

            print(f"\n[{i + 1}/{total}] {org_name}")
            print(f"      URL: {url[:60]}..." if len(url) > 60 else f"      URL: {url}")

            if not url or url == "Not Found":
                print("      → Skipped: No valid URL")
                errors.append({
                    "index": i,
                    "name": org_name,
                    "reason": "No valid Manage URL",
                    "timestamp": datetime.now().isoformat()
                })
                # SAVE CHECKPOINT IMMEDIATELY
                save_checkpoint(processed_data, i + 1, errors, total)
                continue

            # Extract full profile
            profile = extract_organization_full_profile(driver, wait, org)

            if profile:
                processed_data.append(profile)
                print(f"      ✓ Extracted successfully")

                # Show summary of what was found
                contacts_count = len(profile.get("Contacts", []))
                locations_count = len(profile.get("Locations", []))
                sectors_count = len(profile.get("NAICS_Sectors", []))
                web_count = len(profile.get("Web_Presence", []))
                activity_count = len(profile.get("OCIP_Activity", []))

                print(f"         Contacts: {contacts_count} | Locations: {locations_count} | "
                      f"Sectors: {sectors_count} | Web: {web_count} | Activity: {activity_count}")
            else:
                errors.append({
                    "index": i,
                    "name": org_name,
                    "url": url,
                    "reason": "Extraction failed",
                    "timestamp": datetime.now().isoformat()
                })
                print(f"      ✗ Extraction failed")

            # SAVE CHECKPOINT IMMEDIATELY AFTER EVERY ORGANIZATION
            save_checkpoint(processed_data, i + 1, errors, total)

            # Progress indicator
            progress = ((i + 1) / total) * 100
            print(f"      [Progress: {progress:.1f}% | Saved: {len(processed_data)} | Errors: {len(errors)}]")

            # Rate limiting
            if (i + 1) % BATCH_SIZE == 0 and i + 1 < total:
                print(f"\n   [Rate limit pause: {BATCH_PAUSE}s...]")
                time.sleep(BATCH_PAUSE)

            # Delay between organizations
            time.sleep(BETWEEN_ORGS_DELAY)

        # ===== FINAL SAVE =====
        print("\n" + "=" * 60)
        print("SCRAPING COMPLETE!")
        print("=" * 60)
        print(f"Total Organizations Processed: {len(processed_data)}")
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

        total_with_contacts = sum(1 for p in processed_data if p.get("Contacts", []))
        total_with_locations = sum(1 for p in processed_data if p.get("Locations", []))
        total_with_sectors = sum(1 for p in processed_data if p.get("NAICS_Sectors", []))
        total_with_languages = sum(1 for p in processed_data if p.get("Languages_Serviced", []))
        total_with_web = sum(1 for p in processed_data if p.get("Web_Presence", []))
        total_with_activity = sum(1 for p in processed_data if p.get("OCIP_Activity", []))
        total_with_annual = sum(1 for p in processed_data if p.get("Annual_Information", {}))

        print(f"   Organizations with Contacts: {total_with_contacts}")
        print(f"   Organizations with Locations: {total_with_locations}")
        print(f"   Organizations with NAICS Sectors: {total_with_sectors}")
        print(f"   Organizations with Languages: {total_with_languages}")
        print(f"   Organizations with Web Presence: {total_with_web}")
        print(f"   Organizations with OCIP Activity: {total_with_activity}")
        print(f"   Organizations with Annual Info: {total_with_annual}")

        # Clean up checkpoint file on full success
        if os.path.exists(CHECKPOINT_FILE) and len(errors) == 0 and len(processed_data) == total:
            os.remove(CHECKPOINT_FILE)
            print("\n✓ Checkpoint file cleaned up (full success)")

    except KeyboardInterrupt:
        print("\n\n⚠ Script interrupted by user")
        current_idx = i + 1 if 'i' in dir() else start_index
        save_checkpoint(processed_data, current_idx, errors, total)
        print(f"   Progress saved to {CHECKPOINT_FILE}")
        print(f"   Processed {len(processed_data)} organizations so far.")

    except Exception as e:
        print(f"\n✗ CRITICAL ERROR: {e}")
        import traceback
        traceback.print_exc()

        # Emergency save
        if processed_data:
            emergency_file = f"emergency_phase6_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
            with open(emergency_file, 'w', encoding='utf-8') as f:
                json.dump(processed_data, f, indent=2)
            print(f"   Emergency backup saved to {emergency_file}")

        # Also save checkpoint
        current_idx = i + 1 if 'i' in dir() else start_index
        save_checkpoint(processed_data, current_idx, errors, total)

    finally:
        print("\n" + "=" * 60)
        print("Script finished. Browser left open for inspection.")
        print(f"Checkpoint file: {CHECKPOINT_FILE}")
        print(f"Error log: {ERROR_LOG_FILE}")
        print("=" * 60)


if __name__ == "__main__":
    main()