# OCIP Portal Data Extraction Project

## Complete Documentation & User Guide

---

## 📋 Table of Contents

1. [Project Overview](#project-overview)
2. [Architecture](#architecture)
3. [Installation Guide](#installation-guide)
4. [Quick Start](#quick-start)
5. [Phase-by-Phase Guide](#phase-by-phase-guide)
6. [Output Files](#output-files)
7. [Configuration Options](#configuration-options)
8. [Troubleshooting](#troubleshooting)
9. [Data Schema Reference](#data-schema-reference)
10. [Best Practices](#best-practices)

---

## Project Overview

### What This Project Does

This project provides a complete automated data extraction pipeline for the **eCampusOntario OCIP Express Portal** (https://www.ocip.express/). It extracts comprehensive information about:

- **Experts** - Academic and industry experts registered in the system
- **Facilities** - Research facilities and labs across institutions
- **Organizations** - Business and industry partners

### Why Six Phases?

The extraction is split into **6 phases** for reliability and efficiency:

| Phase | Type | Description |
|-------|------|-------------|
| **Phase 1** | Experts Metadata | Collects basic expert info + links from all institutions |
| **Phase 2** | Experts Details | Visits each expert's page to extract full profile |
| **Phase 3** | Facilities Metadata | Collects basic facility info + links from all institutions |
| **Phase 4** | Facilities Details | Visits each facility's page to extract full profile |
| **Phase 5** | Organizations Metadata | Collects basic organization info + links from table |
| **Phase 6** | Organizations Details | Visits each organization's page to extract full profile |

### Key Features

- ✅ **Checkpoint/Resume System** - Never lose progress; resume after interruptions
- ✅ **Real-time Data Access** - View collected data while scripts are running
- ✅ **Rate Limiting** - Built-in delays to avoid overwhelming the server
- ✅ **Error Logging** - Detailed logs of any failed extractions
- ✅ **Multiple Output Formats** - JSON and Excel exports
- ✅ **Duplicate Detection** - Prevents duplicate records when resuming

---

## Architecture

### System Flow Diagram

```
┌─────────────────────────────────────────────────────────────────────┐
│                        OCIP EXTRACTION PIPELINE                      │
└─────────────────────────────────────────────────────────────────────┘

     EXPERTS                  FACILITIES               ORGANIZATIONS
        │                         │                          │
        ▼                         ▼                          ▼
┌───────────────┐        ┌───────────────┐         ┌───────────────┐
│   PHASE 1     │        │   PHASE 3     │         │   PHASE 5     │
│   Metadata    │        │   Metadata    │         │   Metadata    │
│   Harvester   │        │   Harvester   │         │   Harvester   │
└───────┬───────┘        └───────┬───────┘         └───────┬───────┘
        │                        │                         │
        ▼                        ▼                         ▼
┌───────────────┐        ┌───────────────┐         ┌───────────────┐
│ experts_      │        │ facilities_   │         │ organizations_│
│ master_list   │        │ master_list   │         │ master_list   │
│ .json         │        │ .json         │         │ .json         │
└───────┬───────┘        └───────┬───────┘         └───────┬───────┘
        │                        │                         │
        ▼                        ▼                         ▼
┌───────────────┐        ┌───────────────┐         ┌───────────────┐
│   PHASE 2     │        │   PHASE 4     │         │   PHASE 6     │
│   Deep        │        │   Deep        │         │   Deep        │
│   Scraper     │        │   Scraper     │         │   Scraper     │
└───────┬───────┘        └───────┬───────┘         └───────┬───────┘
        │                        │                         │
        ▼                        ▼                         ▼
┌───────────────┐        ┌───────────────┐         ┌───────────────┐
│ experts_      │        │ facilities_   │         │ organizations_│
│ full_details  │        │ full_details  │         │ full_details  │
│ .json         │        │ .json         │         │ .json         │
└───────────────┘        └───────────────┘         └───────────────┘
```

### File Structure

```
OCIP-Extraction/
│
├── 📄 README.md                    # This documentation
├── 📄 requirements.txt             # Python dependencies
├── 📄 setup.bat                    # Windows setup script
│
├── 🐍 phase1_experts_metadata.py   # Expert list harvester
├── 🐍 phase2_experts_details.py    # Expert deep scraper
├── 🐍 phase3_facilities_metadata.py # Facility list harvester
├── 🐍 phase4_facilities_details.py # Facility deep scraper
├── 🐍 phase5_organizations_metadata.py # Organization list harvester
├── 🐍 phase6_organizations_details.py  # Organization deep scraper
│
├── 📂 output/                      # Generated data files
│   ├── experts_master_list.json
│   ├── experts_master_list.xlsx
│   ├── experts_full_details.json
│   ├── facilities_master_list.json
│   ├── facilities_master_list.xlsx
│   ├── facilities_full_details.json
│   ├── organizations_master_list.json
│   ├── organizations_master_list.xlsx
│   └── organizations_full_details.json
│
└── 📂 checkpoints/                 # Progress checkpoints
    ├── checkpoint_progress.json
    ├── phase2_checkpoint.json
    ├── phase3_checkpoint.json
    ├── phase4_checkpoint.json
    ├── organizations_checkpoint.json
    └── phase6_checkpoint.json
```

---

## Installation Guide

### Prerequisites

- **Python 3.9+** installed on your system
- **Google Chrome** browser installed
- **ChromeDriver** (automatically managed by webdriver-manager)
- Valid **OCIP Portal credentials**

### Step-by-Step Installation

#### Option 1: Automated Setup (Windows)

1. **Download/Clone the project** to your local machine

2. **Run the setup script**:
   ```
   Double-click setup.bat
   ```
   
   Or from command prompt:
   ```cmd
   cd path\to\OCIP-Extraction
   setup.bat
   ```

3. **Wait for completion** - The script will:
   - Create a virtual environment (`venv`)
   - Activate the environment
   - Upgrade pip
   - Install all required packages

#### Option 2: Manual Setup (All Platforms)

1. **Open terminal/command prompt** in the project directory

2. **Create virtual environment**:
   ```bash
   python -m venv venv
   ```

3. **Activate virtual environment**:
   
   - **Windows**:
     ```cmd
     venv\Scripts\activate
     ```
   
   - **macOS/Linux**:
     ```bash
     source venv/bin/activate
     ```

4. **Install dependencies**:
   ```bash
   pip install --upgrade pip
   pip install -r requirements.txt
   ```

### Verify Installation

```bash
python -c "import selenium; import pandas; print('All packages installed successfully!')"
```

---

## Quick Start

### Running Your First Extraction

1. **Activate the virtual environment** (if not already active):
   ```cmd
   venv\Scripts\activate
   ```

2. **Run Phase 1** (Experts Metadata):
   ```bash
   python phase1_experts_metadata.py
   ```

3. **Log in manually** when the browser opens:
   - Navigate to the OCIP portal
   - Enter your credentials
   - Navigate to the **Experts Dashboard**
   - Press **ENTER** in the terminal when ready

4. **Let it run** - The script will:
   - Iterate through all institutions
   - Collect expert metadata
   - Save progress automatically

5. **Continue with Phase 2**:
   ```bash
   python phase2_experts_details.py
   ```

### Execution Order

**Always run phases in order within each category:**

```
EXPERTS:        Phase 1 → Phase 2
FACILITIES:     Phase 3 → Phase 4
ORGANIZATIONS:  Phase 5 → Phase 6
```

You can run the three categories in any order or in parallel (using separate terminals).

---

## Phase-by-Phase Guide

### Phase 1: Expert Metadata Harvester

**Purpose**: Collect basic information about all experts across all institutions.

**How it works**:
1. Opens the Experts Dashboard
2. Iterates through the institution dropdown
3. For each institution, scrapes the expert table
4. Handles pagination automatically
5. Saves results to JSON and Excel

**Command**:
```bash
python phase1_experts_metadata.py
```

**Required Input**: Manual login to OCIP portal

**Output Files**:
- `experts_master_list.json`
- `experts_master_list.xlsx`
- `checkpoint_progress.json` (progress tracking)

**Data Collected**:
| Field | Description |
|-------|-------------|
| Institution | The HEI the expert belongs to |
| Name | Expert's full name |
| Expert_Type | Type classification |
| Position | Job title/position |
| Facility | Associated facility |
| Expert_ID | Unique identifier |
| Manage_URL | Link to full profile |
| Profile_URL | Public profile link |

---

### Phase 2: Expert Deep Scraper

**Purpose**: Visit each expert's detail page and extract comprehensive profile data.

**How it works**:
1. Loads the master list from Phase 1
2. Visits each expert's Manage_URL
3. Expands all accordion sections
4. Extracts data from each section
5. Saves detailed profiles

**Command**:
```bash
python phase2_experts_details.py
```

**Required Input**: 
- `experts_master_list.json` from Phase 1
- Manual login to OCIP portal

**Output Files**:
- `experts_full_details.json`
- `phase2_checkpoint.json`
- `phase2_errors.json`

**Sections Extracted**:
1. General Information
2. Details (Biography, Description)
3. Expert Demographics
4. Expertise (SRED codes, disciplines)
5. Price & Availability
6. Facility Affiliation
7. Web Presence
8. OCIP Activity
9. Audit Trail

---

### Phase 3: Facility Metadata Harvester

**Purpose**: Collect basic information about all facilities across all institutions.

**How it works**:
1. Opens the Facilities Admin page
2. Iterates through the institution dropdown
3. For each institution, scrapes the facility table
4. Handles pagination automatically

**Command**:
```bash
python phase3_facilities_metadata.py
```

**URL**: `https://www.ocip.express/FacilityAdmin/Index`

**Output Files**:
- `facilities_master_list.json`
- `facilities_master_list.xlsx`
- `phase3_checkpoint.json`

**Data Collected**:
| Field | Description |
|-------|-------------|
| Institution | The HEI the facility belongs to |
| Facility_Name | Name of the facility |
| Facility_ID | Unique identifier |
| Type | Facility type classification |
| Manage_URL | Link to full details |

---

### Phase 4: Facility Deep Scraper

**Purpose**: Visit each facility's detail page and extract comprehensive data.

**Command**:
```bash
python phase4_facilities_details.py
```

**Required Input**: `facilities_master_list.json` from Phase 3

**Output Files**:
- `facilities_full_details.json`
- `phase4_checkpoint.json`
- `phase4_errors.json`

**Sections Extracted** (12 total):
1. General Information
2. Academic Unit Details
3. Provinces Served
4. Activities Offered
5. Sectors Served
6. Contacts
7. Locations
8. Facility Descriptors
9. Languages Serviced
10. Web Presence
11. OCIP Activity
12. Audit Trail

---

### Phase 5: Organization Metadata Harvester

**Purpose**: Collect basic information about all organizations from the single paginated table.

**How it works**:
1. Opens the Business Admin page
2. Scrapes the organization table
3. Handles pagination (no dropdown needed)
4. Saves after every page

**Command**:
```bash
python phase5_organizations_metadata.py
```

**URL**: `https://www.ocip.express/BusinessAdmin/Index`

**Output Files**:
- `organizations_master_list.json`
- `organizations_master_list.xlsx`
- `organizations_checkpoint.json`

**Data Collected**:
| Field | Description |
|-------|-------------|
| Organization_Name | Name of the organization |
| Provinces | Provinces where organization operates |
| Sectors | Industry sectors |
| Requests | Has active requests (Yes/No) |
| Projects | Has active projects (Yes/No) |
| Enabled | Account is enabled (Yes/No) |
| Manage_URL | Link to full details |

**Column Mapping** (for reference):
```
td[4] (index 3) = Organization Name
td[5] (index 4) = Provinces
td[6] (index 5) = Sectors
td[7] (index 6) = Requests?
td[8] (index 7) = Projects?
td[9] (index 8) = Enabled
td[10] (index 9) = Actions (Manage Link)
```

---

### Phase 6: Organization Deep Scraper

**Purpose**: Visit each organization's detail page and extract comprehensive data.

**Command**:
```bash
python phase6_organizations_details.py
```

**Required Input**: `organizations_master_list.json` from Phase 5

**Output Files**:
- `organizations_full_details.json`
- `phase6_checkpoint.json`
- `phase6_errors.json`

**Sections Extracted** (10 total):
1. General Information
2. Organization Information
3. Annual Information
4. NAICS Sectors
5. Contacts
6. Locations
7. Languages Serviced
8. Web Presence
9. OCIP Activity
10. Audit Trail

---

## Output Files

### File Formats

| Format | Use Case |
|--------|----------|
| **JSON** | Primary data format; preserves nested structures; ideal for programmatic access |
| **Excel** | Flat table format; ideal for manual review and basic analysis |

### JSON Structure Examples

#### Master List (Metadata) Format
```json
[
  {
    "Organization_Name": "Example Corp",
    "Provinces": "Ontario",
    "Sectors": "Manufacturing",
    "Requests": "Yes",
    "Projects": "No",
    "Enabled": "Yes",
    "Manage_URL": "https://www.ocip.express/...",
    "Scraped_At": "2024-01-15T10:30:00"
  }
]
```

#### Full Details Format
```json
[
  {
    "Meta": {
      "Source_URL": "https://...",
      "Scraped_At": "2024-01-15T10:30:00",
      "Organization_Name_From_List": "Example Corp"
    },
    "General_Information": {
      "Name": "Example Corporation",
      "Status": "Active"
    },
    "Contacts": [
      {
        "Name": "John Doe",
        "Email": "john@example.com",
        "Phone": "555-1234"
      }
    ],
    "Locations": [...],
    "NAICS_Sectors": [...],
    ...
  }
]
```

### Checkpoint Files

Checkpoint files are **JSON formatted** and can be opened anytime during execution:

```json
{
  "timestamp": "2024-01-15T10:30:00",
  "current_index": 150,
  "total_organizations": 500,
  "organizations_processed": 148,
  "errors_count": 2,
  "progress_percent": 30.0,
  "data": [...]
}
```

---

## Configuration Options

### Timing Configuration

Each script has configurable timing parameters at the top:

```python
# Timing Configuration
PAGE_LOAD_WAIT = 2.0       # Seconds to wait after page loads
PAGINATION_WAIT = 1.5      # Seconds to wait after clicking next page
ACCORDION_EXPAND_WAIT = 0.5 # Seconds to wait after expanding accordion
BETWEEN_ITEMS_DELAY = 1.0  # Seconds between processing items
REQUEST_DELAY = 0.3        # Seconds between accordion section extractions

# Rate Limiting
BATCH_SIZE = 50            # Process this many items before pausing
BATCH_PAUSE = 10           # Seconds to pause between batches
```

### Adjusting for Slow Connections

If you have a slow internet connection, increase these values:

```python
PAGE_LOAD_WAIT = 4.0
PAGINATION_WAIT = 3.0
LOADING_MASK_TIMEOUT = 20
```

### Adjusting for Fast Connections

If you have a fast connection and want to speed up:

```python
PAGE_LOAD_WAIT = 1.0
PAGINATION_WAIT = 0.8
BETWEEN_ITEMS_DELAY = 0.5
```

⚠️ **Warning**: Setting values too low may cause missed data or errors.

---

## Troubleshooting

### Common Issues and Solutions

#### 1. "ChromeDriver not found" Error

**Solution**: The `webdriver-manager` package should handle this automatically. If not:
```bash
pip install --upgrade webdriver-manager
```

#### 2. "Element not found" or Timeout Errors

**Causes**:
- Page hasn't loaded completely
- Website structure has changed
- Network issues

**Solutions**:
- Increase `PAGE_LOAD_WAIT` and `LOADING_MASK_TIMEOUT`
- Check your internet connection
- Verify you're logged in correctly

#### 3. Script Stops Unexpectedly

**Solution**: Use the checkpoint system to resume:
```
When prompted "Resume from checkpoint? (y/n):", enter 'y'
```

#### 4. Empty Data in Output

**Causes**:
- Not logged in properly
- Not on the correct page when pressing ENTER
- Institution has no data

**Solutions**:
- Ensure you're fully logged in
- Navigate to the correct dashboard before pressing ENTER
- Check the checkpoint file for partial data

#### 5. "Stale Element Reference" Errors

**Cause**: Page updated while script was reading it

**Solution**: These are usually handled automatically. If persistent, increase delay values.

#### 6. Browser Closes Unexpectedly

**Note**: Scripts are configured with `detach=True` to keep the browser open. If it closes:
- Check for Python errors in the terminal
- The browser may have crashed; restart the script

### Checking Progress During Execution

You can open checkpoint files **while the script is running**:

```bash
# Windows
type organizations_checkpoint.json

# Or open in any text editor
notepad organizations_checkpoint.json
```

The file is written and flushed to disk after every item/page.

---

## Data Schema Reference

### Expert Full Details Schema

```
Expert Profile
├── Meta
│   ├── Source_URL
│   ├── Scraped_At
│   ├── Institution
│   ├── Name_From_List
│   ├── Expert_ID
│   └── Profile_URL
│
├── General_Information
│   ├── Academic_Unit
│   ├── IsLinkedToUser (Yes/No)
│   ├── Enabled (Yes/No)
│   ├── Email
│   ├── Phone
│   ├── Reputation_Score
│   └── Photo_URL
│
├── Details
│   ├── ProfileDescription
│   └── [Other fields]
│
├── Expert_Demographics
│   └── [Demographic fields]
│
├── Expertise (Array)
│   └── {SRED_Code, Area, Discipline, Field}
│
├── Price_Availability
│   ├── Daily_Rate
│   ├── Can_Initiate_Innovation_Challenge
│   ├── Available_for_Scoping
│   ├── Available_for_Projects
│   └── Can_be_Principal_Investigator
│
├── Facility_Affiliation (Array)
│   └── {Facility_Name, Is_Primary_Facility}
│
├── Web_Presence (Array)
│   └── {Name, Type, URL}
│
├── OCIP_Activity (Array)
│   └── {Project_Name, Project_URL, Type, Organization, Current_Status}
│
└── Audit_Trail
    ├── Created_By
    ├── Created_Date
    ├── Modified_By
    └── Modified_Date
```

### Facility Full Details Schema

```
Facility Profile
├── Meta
│   ├── Source_URL
│   ├── Scraped_At
│   ├── Institution
│   ├── Facility_Name_From_List
│   ├── Facility_ID
│   └── Type_From_List
│
├── General_Information
├── Academic_Unit_Details
├── Provinces_Served (Array)
├── Activities_Offered (Array)
├── Sectors_Served (Array)
├── Contacts (Array)
├── Locations (Array)
├── Facility_Descriptors
├── Languages_Serviced (Array)
├── Web_Presence (Array)
├── OCIP_Activity (Array)
└── Audit_Trail
```

### Organization Full Details Schema

```
Organization Profile
├── Meta
│   ├── Source_URL
│   ├── Scraped_At
│   ├── Organization_Name_From_List
│   ├── Provinces_From_List
│   └── Sectors_From_List
│
├── General_Information
├── Organization_Information
├── Annual_Information
├── NAICS_Sectors (Array)
├── Contacts (Array)
├── Locations (Array)
├── Languages_Serviced (Array)
├── Web_Presence (Array)
├── OCIP_Activity (Array)
└── Audit_Trail
```

---

## Best Practices

### Before Running

1. ✅ **Ensure stable internet connection**
2. ✅ **Close unnecessary browser tabs** to free up memory
3. ✅ **Have valid OCIP credentials** ready
4. ✅ **Check available disk space** for output files

### During Execution

1. ✅ **Don't interact with the browser** while scripts are running
2. ✅ **Monitor the terminal** for progress and errors
3. ✅ **Let the script handle pagination** - don't click anything
4. ✅ **Check checkpoint files** periodically if running long jobs

### After Completion

1. ✅ **Verify output files** exist and contain data
2. ✅ **Review error logs** for any failed extractions
3. ✅ **Backup important data** before re-running scripts
4. ✅ **Clean up checkpoint files** if extraction is complete

### Performance Tips

| Tip | Benefit |
|-----|---------|
| Run during off-peak hours | Faster page loads, fewer timeouts |
| Use wired internet connection | More stable than WiFi |
| Close other applications | More memory for Chrome |
| Run phases sequentially | Prevents resource conflicts |

### Data Integrity

- **Always run metadata phase before details phase**
- **Don't modify master list files** between phases
- **Use checkpoint resume** instead of restarting from scratch
- **Keep backup copies** of completed extractions

---

## Support & Maintenance

### Updating the Scripts

If the OCIP portal structure changes:

1. Inspect the new HTML structure using browser DevTools
2. Update the CSS selectors in the relevant phase script
3. Test on a small subset before full extraction

### Common Selector Locations

| Element | Typical Selector |
|---------|-----------------|
| Data rows | `tr.k-master-row` |
| Pagination info | `span.k-pager-info.k-label` |
| Next page button | `a.k-pager-nav[aria-label='Go to the next page']` |
| Accordion headers | `li.k-panelbar-header` |
| Loading mask | `.k-loading-mask` |
| Dropdown listbox | `ul[id$='listbox']` |

---

## License & Disclaimer

This tool is provided for **authorized use only**. Ensure you have proper permission to access and extract data from the OCIP Express portal. The authors are not responsible for any misuse of this tool or violation of terms of service.

---

**Last Updated**: January 2024  
**Version**: 2.0  
**Author**: AI Assistant
```

---

And here is a **Quick Reference Card** as a separate, shorter file:

```markdown
# OCIP Extraction - Quick Reference Card

## 🚀 Setup (One-Time)

```cmd
# Windows - Double-click setup.bat
# OR run manually:
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

## 📋 Execution Order

```
EXPERTS:        python phase1_experts_metadata.py → python phase2_experts_details.py
FACILITIES:     python phase3_facilities_metadata.py → python phase4_facilities_details.py  
ORGANIZATIONS:  python phase5_organizations_metadata.py → python phase6_organizations_details.py
```

## 🔑 Key Steps for Each Phase

1. Run the script
2. Browser opens → **Log in manually**
3. Navigate to correct dashboard page
4. Press **ENTER** in terminal
5. Wait for completion

## 📁 Output Files

| Phase | Output |
|-------|--------|
| 1 | `experts_master_list.json` |
| 2 | `experts_full_details.json` |
| 3 | `facilities_master_list.json` |
| 4 | `facilities_full_details.json` |
| 5 | `organizations_master_list.json` |
| 6 | `organizations_full_details.json` |

## 💾 Resume from Checkpoint

When prompted: `Resume from checkpoint? (y/n):` → Type `y`

## 👀 View Progress (While Running)

```cmd
type phase6_checkpoint.json
```

## ⚠️ If Something Goes Wrong

1. Check terminal for error messages
2. Look at the error log file (e.g., `phase6_errors.json`)
3. Resume from checkpoint
4. If persistent, increase timing values in script

## ⏱️ Timing Adjustments

Edit at top of each script:
```python
PAGE_LOAD_WAIT = 2.0      # Increase if pages load slowly
PAGINATION_WAIT = 1.5     # Increase if pagination fails
BATCH_PAUSE = 10          # Increase to be gentler on server
```
