"""
OCIP Portal Data Extraction - Main Controller
==============================================
Central orchestration script for all 6 extraction phases.

Features:
- Single login session shared across phases
- Interactive menu system
- Automatic folder organization
- Progress tracking dashboard
- Custom execution modes (single, sequential, full pipeline)

Author: AI Assistant
Version: 1.0
"""

import os
import sys
import json
import time
import shutil
from datetime import datetime
from pathlib import Path

# Selenium imports
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# ==========================================
# DIRECTORY CONFIGURATION
# ==========================================
BASE_DIR = Path(__file__).parent
OUTPUT_DIR = BASE_DIR / "output"
CHECKPOINT_DIR = BASE_DIR / "checkpoints"
LOGS_DIR = BASE_DIR / "logs"

# Create directories if they don't exist
OUTPUT_DIR.mkdir(exist_ok=True)
CHECKPOINT_DIR.mkdir(exist_ok=True)
LOGS_DIR.mkdir(exist_ok=True)

# ==========================================
# FILE PATH CONFIGURATION
# ==========================================
FILE_PATHS = {
    # Phase 1: Experts Metadata
    "phase1": {
        "output_json": OUTPUT_DIR / "experts_master_list.json",
        "output_excel": OUTPUT_DIR / "experts_master_list.xlsx",
        "checkpoint": CHECKPOINT_DIR / "phase1_experts_checkpoint.json",
    },
    # Phase 2: Experts Details
    "phase2": {
        "input": OUTPUT_DIR / "experts_master_list.json",
        "output_json": OUTPUT_DIR / "experts_full_details.json",
        "checkpoint": CHECKPOINT_DIR / "phase2_experts_checkpoint.json",
        "errors": LOGS_DIR / "phase2_experts_errors.json",
    },
    # Phase 3: Facilities Metadata
    "phase3": {
        "output_json": OUTPUT_DIR / "facilities_master_list.json",
        "output_excel": OUTPUT_DIR / "facilities_master_list.xlsx",
        "checkpoint": CHECKPOINT_DIR / "phase3_facilities_checkpoint.json",
    },
    # Phase 4: Facilities Details
    "phase4": {
        "input": OUTPUT_DIR / "facilities_master_list.json",
        "output_json": OUTPUT_DIR / "facilities_full_details.json",
        "checkpoint": CHECKPOINT_DIR / "phase4_facilities_checkpoint.json",
        "errors": LOGS_DIR / "phase4_facilities_errors.json",
    },
    # Phase 5: Organizations Metadata
    "phase5": {
        "output_json": OUTPUT_DIR / "organizations_master_list.json",
        "output_excel": OUTPUT_DIR / "organizations_master_list.xlsx",
        "checkpoint": CHECKPOINT_DIR / "phase5_organizations_checkpoint.json",
    },
    # Phase 6: Organizations Details
    "phase6": {
        "input": OUTPUT_DIR / "organizations_master_list.json",
        "output_json": OUTPUT_DIR / "organizations_full_details.json",
        "checkpoint": CHECKPOINT_DIR / "phase6_organizations_checkpoint.json",
        "errors": LOGS_DIR / "phase6_organizations_errors.json",
    },
}

# ==========================================
# URL CONFIGURATION
# ==========================================
URLS = {
    "login": "https://www.ocip.express/",
    "experts": "https://www.ocip.express/ExpertAdmin/Index",
    "facilities": "https://www.ocip.express/FacilityAdmin/Index",
    "organizations": "https://www.ocip.express/BusinessAdmin/Index",
}

# ==========================================
# PHASE METADATA
# ==========================================
PHASE_INFO = {
    1: {
        "name": "Experts Metadata",
        "description": "Harvest expert list from all institutions",
        "category": "Experts",
        "type": "Metadata",
        "url": URLS["experts"],
        "depends_on": None,
    },
    2: {
        "name": "Experts Details",
        "description": "Extract full profiles for each expert",
        "category": "Experts",
        "type": "Details",
        "url": None,  # Uses URLs from Phase 1 output
        "depends_on": 1,
    },
    3: {
        "name": "Facilities Metadata",
        "description": "Harvest facility list from all institutions",
        "category": "Facilities",
        "type": "Metadata",
        "url": URLS["facilities"],
        "depends_on": None,
    },
    4: {
        "name": "Facilities Details",
        "description": "Extract full profiles for each facility",
        "category": "Facilities",
        "type": "Details",
        "url": None,
        "depends_on": 3,
    },
    5: {
        "name": "Organizations Metadata",
        "description": "Harvest organization list from table",
        "category": "Organizations",
        "type": "Metadata",
        "url": URLS["organizations"],
        "depends_on": None,
    },
    6: {
        "name": "Organizations Details",
        "description": "Extract full profiles for each organization",
        "category": "Organizations",
        "type": "Details",
        "url": None,
        "depends_on": 5,
    },
}


# ==========================================
# DISPLAY UTILITIES
# ==========================================
def clear_screen():
    """Clear terminal screen."""
    os.system('cls' if os.name == 'nt' else 'clear')


def print_header():
    """Print application header."""
    print("\n" + "=" * 70)
    print("  ██████╗  ██████╗██╗██████╗     ███████╗██╗  ██╗████████╗")
    print(" ██╔═══██╗██╔════╝██║██╔══██╗    ██╔════╝╚██╗██╔╝╚══██╔══╝")
    print(" ██║   ██║██║     ██║██████╔╝    █████╗   ╚███╔╝    ██║   ")
    print(" ██║   ██║██║     ██║██╔═══╝     ██╔══╝   ██╔██╗    ██║   ")
    print(" ╚██████╔╝╚██████╗██║██║         ███████╗██╔╝ ██╗   ██║   ")
    print("  ╚═════╝  ╚═════╝╚═╝╚═╝         ╚══════╝╚═╝  ╚═╝   ╚═╝   ")
    print("=" * 70)
    print("  OCIP Portal Data Extraction Pipeline - Main Controller")
    print("=" * 70)


def print_divider(char="-", length=70):
    """Print a divider line."""
    print(char * length)


def format_file_size(size_bytes):
    """Format file size in human-readable format."""
    if size_bytes == 0:
        return "0 B"
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_bytes < 1024:
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024
    return f"{size_bytes:.1f} TB"


def get_file_info(filepath):
    """Get file information if exists."""
    path = Path(filepath)
    if path.exists():
        stat = path.stat()
        size = format_file_size(stat.st_size)
        modified = datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M")
        return {"exists": True, "size": size, "modified": modified}
    return {"exists": False, "size": "-", "modified": "-"}


def get_checkpoint_progress(checkpoint_path):
    """Read checkpoint and return progress info."""
    try:
        with open(checkpoint_path, 'r', encoding='utf-8') as f:
            data = json.load(f)

        # Extract common fields
        processed = data.get('experts_processed') or data.get('facilities_processed') or \
                    data.get('organizations_processed') or data.get('organizations_collected') or 0
        total = data.get('total_institutions') or data.get('total_organizations') or \
                data.get('total_items_in_table') or 0
        current_idx = data.get('current_index') or data.get('last_page_scraped') or 0

        if total > 0:
            percent = (current_idx / total) * 100
        else:
            percent = 0

        return {
            "processed": processed,
            "total": total,
            "percent": percent,
            "timestamp": data.get('timestamp', 'Unknown')
        }
    except:
        return None


def count_json_records(filepath):
    """Count records in a JSON file."""
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            data = json.load(f)
        if isinstance(data, list):
            return len(data)
        elif isinstance(data, dict) and 'data' in data:
            return len(data['data'])
    except:
        pass
    return 0


# ==========================================
# STATUS DASHBOARD
# ==========================================
def show_status_dashboard():
    """Display comprehensive status of all phases."""
    clear_screen()
    print_header()
    print("\n  📊 EXTRACTION STATUS DASHBOARD")
    print_divider("=")

    print("\n  {:^8} │ {:^22} │ {:^10} │ {:^12} │ {:^15}".format(
        "Phase", "Name", "Status", "Records", "Last Modified"
    ))
    print("  " + "-" * 8 + "─┼─" + "-" * 22 + "─┼─" + "-" * 10 + "─┼─" + "-" * 12 + "─┼─" + "-" * 15)

    for phase_num, info in PHASE_INFO.items():
        phase_key = f"phase{phase_num}"
        paths = FILE_PATHS[phase_key]

        # Check output file
        output_path = paths.get("output_json")
        file_info = get_file_info(output_path)

        # Check checkpoint
        checkpoint_path = paths.get("checkpoint")
        checkpoint_progress = get_checkpoint_progress(checkpoint_path) if Path(checkpoint_path).exists() else None

        # Determine status
        if file_info["exists"]:
            records = count_json_records(output_path)
            if checkpoint_progress and checkpoint_progress["percent"] < 100:
                status = "🔄 Partial"
            else:
                status = "✅ Complete"
        elif checkpoint_progress:
            status = "⏸️ Paused"
            records = checkpoint_progress["processed"]
        else:
            status = "⬜ Pending"
            records = 0

        print("  {:^8} │ {:<22} │ {:^10} │ {:>12} │ {:^15}".format(
            f"[{phase_num}]",
            info["name"][:22],
            status,
            f"{records:,}" if records else "-",
            file_info["modified"] if file_info["exists"] else "-"
        ))

    print_divider("=")

    # Show folder sizes
    print("\n  📁 FOLDER SUMMARY")
    print_divider()

    output_size = sum(f.stat().st_size for f in OUTPUT_DIR.glob("*") if f.is_file())
    checkpoint_size = sum(f.stat().st_size for f in CHECKPOINT_DIR.glob("*") if f.is_file())
    logs_size = sum(f.stat().st_size for f in LOGS_DIR.glob("*") if f.is_file())

    print(f"  Output folder:     {format_file_size(output_size):>10}  ({len(list(OUTPUT_DIR.glob('*')))} files)")
    print(
        f"  Checkpoints folder:{format_file_size(checkpoint_size):>10}  ({len(list(CHECKPOINT_DIR.glob('*')))} files)")
    print(f"  Logs folder:       {format_file_size(logs_size):>10}  ({len(list(LOGS_DIR.glob('*')))} files)")

    print()
    input("  Press ENTER to return to main menu...")


# ==========================================
# BROWSER SESSION MANAGEMENT
# ==========================================
class BrowserSession:
    """Manages a single browser session across multiple phases."""

    def __init__(self):
        self.driver = None
        self.wait = None
        self.is_logged_in = False
        self.current_page = None

    def start(self):
        """Initialize the browser."""
        if self.driver is not None:
            return  # Already started

        print("\n  🌐 Starting browser...")

        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option("useAutomationExtension", False)
        options.add_experimental_option("detach", True)

        self.driver = webdriver.Chrome(options=options)
        self.driver.execute_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        )
        self.wait = WebDriverWait(self.driver, 20)

        print("  ✅ Browser started successfully")

    def login(self, force=False):
        """Handle login process."""
        if self.is_logged_in and not force:
            print("  ✅ Already logged in")
            return True

        self.start()

        print(f"\n  🔐 Navigating to login page: {URLS['login']}")
        self.driver.get(URLS["login"])

        print_divider()
        print("  Please log in to the OCIP portal manually.")
        print("  Complete the login process in the browser window.")
        print_divider()

        input("\n  >>> Press ENTER here once you are logged in...")

        self.is_logged_in = True
        print("  ✅ Login confirmed")
        return True

    def navigate_to(self, url, page_name="page"):
        """Navigate to a specific URL."""
        if not self.driver:
            self.start()

        print(f"\n  🔗 Navigating to {page_name}...")
        self.driver.get(url)
        time.sleep(2)  # Allow page to load
        self.current_page = url
        print(f"  ✅ Arrived at {page_name}")

    def close(self):
        """Close the browser."""
        if self.driver:
            print("\n  🔒 Closing browser...")
            try:
                self.driver.quit()
            except:
                pass
            self.driver = None
            self.wait = None
            self.is_logged_in = False
            self.current_page = None
            print("  ✅ Browser closed")

    def is_active(self):
        """Check if browser session is active."""
        if not self.driver:
            return False
        try:
            _ = self.driver.current_url
            return True
        except:
            self.driver = None
            return False


# Global browser session
browser = BrowserSession()


# ==========================================
# PHASE EXECUTION IMPORTS
# ==========================================
# Import phase modules dynamically to avoid circular imports
def import_phase_module(phase_num):
    """Dynamically import a phase module."""
    module_names = {
        1: "phase1_experts_metadata",
        2: "phase2_experts_details",
        3: "phase3_facilities_metadata",
        4: "phase4_facilities_details",
        5: "phase5_organizations_metadata",
        6: "phase6_organizations_details",
    }

    module_name = module_names.get(phase_num)
    if not module_name:
        raise ValueError(f"Invalid phase number: {phase_num}")

    # Check if module file exists
    module_path = BASE_DIR / f"{module_name}.py"
    if not module_path.exists():
        raise FileNotFoundError(f"Module file not found: {module_path}")

    # Import the module
    import importlib.util
    spec = importlib.util.spec_from_file_location(module_name, module_path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)

    return module


# ==========================================
# PHASE EXECUTION WRAPPERS
# ==========================================
def check_phase_dependency(phase_num):
    """Check if phase dependencies are satisfied."""
    info = PHASE_INFO[phase_num]
    depends_on = info.get("depends_on")

    if depends_on is None:
        return True, None

    # Check if dependency output exists
    dep_key = f"phase{depends_on}"
    dep_output = FILE_PATHS[dep_key]["output_json"]

    if not Path(dep_output).exists():
        return False, depends_on

    # Check if it has records
    records = count_json_records(dep_output)
    if records == 0:
        return False, depends_on

    return True, None


def run_phase(phase_num, session=None):
    """Execute a specific phase."""
    info = PHASE_INFO[phase_num]
    phase_key = f"phase{phase_num}"
    paths = FILE_PATHS[phase_key]

    clear_screen()
    print_header()
    print(f"\n  🚀 PHASE {phase_num}: {info['name'].upper()}")
    print(f"     {info['description']}")
    print_divider("=")

    # Check dependencies
    dep_ok, missing_phase = check_phase_dependency(phase_num)
    if not dep_ok:
        print(f"\n  ❌ ERROR: Phase {missing_phase} must be completed first!")
        print(f"     Run Phase {missing_phase} to generate required input data.")
        input("\n  Press ENTER to return to menu...")
        return False

    # Check for existing checkpoint
    checkpoint_path = paths.get("checkpoint")
    if Path(checkpoint_path).exists():
        progress = get_checkpoint_progress(checkpoint_path)
        if progress:
            print(f"\n  ⚠️  Found existing checkpoint:")
            print(f"      Timestamp: {progress['timestamp']}")
            print(f"      Progress: {progress['processed']} processed ({progress['percent']:.1f}%)")

            choice = input("\n  Resume from checkpoint? (y/n/c to cancel): ").strip().lower()
            if choice == 'c':
                return False
            resume = (choice == 'y')
        else:
            resume = False
    else:
        resume = False

    # Use provided session or global browser
    use_session = session if session else browser

    # Ensure logged in
    if not use_session.is_active() or not use_session.is_logged_in:
        use_session.login()

    # Navigate to appropriate page for metadata phases
    if info["url"]:
        use_session.navigate_to(info["url"], info["name"])

    print("\n  Starting extraction...")
    print_divider()

    try:
        # Import and run the phase module
        module = import_phase_module(phase_num)

        # Inject our configuration
        inject_paths_to_module(module, phase_num)

        # Call the module's run function with our driver
        if hasattr(module, 'run_with_driver'):
            # Preferred method: pass driver directly
            result = module.run_with_driver(
                driver=use_session.driver,
                wait=use_session.wait,
                resume=resume
            )
        elif hasattr(module, 'main'):
            # Fallback: module manages its own driver
            print("  ⚠️  Module will manage its own browser session")
            result = module.main()
        else:
            print("  ❌ Module has no executable function")
            return False

        print_divider()
        print(f"\n  ✅ Phase {phase_num} completed successfully!")

        # Show output summary
        output_path = paths.get("output_json")
        if Path(output_path).exists():
            records = count_json_records(output_path)
            file_info = get_file_info(output_path)
            print(f"     Output: {output_path.name}")
            print(f"     Records: {records:,}")
            print(f"     Size: {file_info['size']}")

        return True

    except FileNotFoundError as e:
        print(f"\n  ❌ Module not found: {e}")
        print("     Make sure all phase scripts are in the same directory.")
        return False
    except Exception as e:
        print(f"\n  ❌ Error during execution: {e}")
        import traceback
        traceback.print_exc()
        return False
    finally:
        input("\n  Press ENTER to continue...")


def inject_paths_to_module(module, phase_num):
    """Inject file paths into a module."""
    phase_key = f"phase{phase_num}"
    paths = FILE_PATHS[phase_key]

    # Common path attributes to set
    path_mappings = {
        "OUTPUT_JSON": "output_json",
        "OUTPUT_EXCEL": "output_excel",
        "INPUT_FILE": "input",
        "CHECKPOINT_FILE": "checkpoint",
        "ERROR_LOG_FILE": "errors",
    }

    for attr, key in path_mappings.items():
        if key in paths and hasattr(module, attr):
            setattr(module, attr, str(paths[key]))


# ==========================================
# EXECUTION MODES
# ==========================================
def run_single_phase():
    """Run a single phase selected by user."""
    clear_screen()
    print_header()
    print("\n  📌 SELECT PHASE TO RUN")
    print_divider("=")

    print("\n  EXPERTS:")
    print("    [1] Phase 1 - Experts Metadata")
    print("    [2] Phase 2 - Experts Details")
    print("\n  FACILITIES:")
    print("    [3] Phase 3 - Facilities Metadata")
    print("    [4] Phase 4 - Facilities Details")
    print("\n  ORGANIZATIONS:")
    print("    [5] Phase 5 - Organizations Metadata")
    print("    [6] Phase 6 - Organizations Details")
    print("\n    [0] Back to Main Menu")

    print_divider()
    choice = input("  Enter phase number: ").strip()

    if choice == '0':
        return

    try:
        phase_num = int(choice)
        if 1 <= phase_num <= 6:
            run_phase(phase_num)
        else:
            print("  Invalid choice. Please enter 1-6.")
            time.sleep(1)
    except ValueError:
        print("  Invalid input. Please enter a number.")
        time.sleep(1)


def run_category_pipeline():
    """Run all phases for a specific category."""
    clear_screen()
    print_header()
    print("\n  📦 SELECT CATEGORY PIPELINE")
    print_divider("=")

    print("\n    [1] Experts Pipeline     (Phase 1 → Phase 2)")
    print("    [2] Facilities Pipeline  (Phase 3 → Phase 4)")
    print("    [3] Organizations Pipeline (Phase 5 → Phase 6)")
    print("\n    [0] Back to Main Menu")

    print_divider()
    choice = input("  Enter choice: ").strip()

    if choice == '0':
        return

    pipelines = {
        '1': [1, 2],
        '2': [3, 4],
        '3': [5, 6],
    }

    if choice not in pipelines:
        print("  Invalid choice.")
        time.sleep(1)
        return

    phases = pipelines[choice]
    category = PHASE_INFO[phases[0]]["category"]

    print(f"\n  🚀 Starting {category} Pipeline...")
    print(f"     Phases to run: {phases}")

    confirm = input("\n  Continue? (y/n): ").strip().lower()
    if confirm != 'y':
        return

    # Ensure login once
    browser.login()

    for phase_num in phases:
        success = run_phase(phase_num, session=browser)
        if not success:
            print(f"\n  ⚠️  Pipeline stopped at Phase {phase_num}")
            break

    print(f"\n  📦 {category} Pipeline completed!")
    input("  Press ENTER to continue...")


def run_full_pipeline():
    """Run all 6 phases in sequence."""
    clear_screen()
    print_header()
    print("\n  🔄 FULL EXTRACTION PIPELINE")
    print_divider("=")

    print("\n  This will run ALL 6 phases in sequence:")
    print("    Phase 1 → Phase 2 → Phase 3 → Phase 4 → Phase 5 → Phase 6")
    print("\n  ⚠️  This may take several hours depending on data volume.")

    confirm = input("\n  Are you sure you want to continue? (yes/no): ").strip().lower()
    if confirm != 'yes':
        return

    # Single login for entire pipeline
    browser.login()

    start_time = datetime.now()
    completed = []
    failed = []

    for phase_num in range(1, 7):
        print(f"\n  {'=' * 50}")
        print(f"  Starting Phase {phase_num} of 6...")
        print(f"  {'=' * 50}")

        success = run_phase(phase_num, session=browser)

        if success:
            completed.append(phase_num)
        else:
            failed.append(phase_num)
            print(f"\n  ⚠️  Phase {phase_num} failed. Continue with remaining? (y/n): ")
            if input().strip().lower() != 'y':
                break

    # Summary
    end_time = datetime.now()
    duration = end_time - start_time

    clear_screen()
    print_header()
    print("\n  📊 FULL PIPELINE SUMMARY")
    print_divider("=")

    print(f"\n  Duration: {duration}")
    print(f"  Completed Phases: {completed}")
    print(f"  Failed Phases: {failed}")

    input("\n  Press ENTER to continue...")


# ==========================================
# UTILITY FUNCTIONS
# ==========================================
def clean_checkpoints():
    """Clean all checkpoint files."""
    clear_screen()
    print_header()
    print("\n  🧹 CLEAN CHECKPOINTS")
    print_divider("=")

    checkpoint_files = list(CHECKPOINT_DIR.glob("*.json"))

    if not checkpoint_files:
        print("\n  No checkpoint files found.")
        input("  Press ENTER to continue...")
        return

    print(f"\n  Found {len(checkpoint_files)} checkpoint file(s):")
    for f in checkpoint_files:
        info = get_file_info(f)
        print(f"    - {f.name} ({info['size']}, modified: {info['modified']})")

    print("\n  ⚠️  WARNING: This will delete all checkpoint files.")
    print("     You will lose the ability to resume incomplete extractions.")

    confirm = input("\n  Type 'DELETE' to confirm: ").strip()

    if confirm == 'DELETE':
        for f in checkpoint_files:
            f.unlink()
        print(f"\n  ✅ Deleted {len(checkpoint_files)} checkpoint file(s).")
    else:
        print("\n  ❌ Operation cancelled.")

    input("  Press ENTER to continue...")


def clean_all_data():
    """Clean all output, checkpoints, and logs."""
    clear_screen()
    print_header()
    print("\n  ⚠️  CLEAN ALL DATA")
    print_divider("=")

    print("\n  This will delete:")
    print(f"    - All files in {OUTPUT_DIR}")
    print(f"    - All files in {CHECKPOINT_DIR}")
    print(f"    - All files in {LOGS_DIR}")

    print("\n  ⚠️  THIS ACTION CANNOT BE UNDONE!")

    confirm = input("\n  Type 'DELETE ALL' to confirm: ").strip()

    if confirm == 'DELETE ALL':
        for folder in [OUTPUT_DIR, CHECKPOINT_DIR, LOGS_DIR]:
            for f in folder.glob("*"):
                if f.is_file():
                    f.unlink()
        print("\n  ✅ All data files deleted.")
    else:
        print("\n  ❌ Operation cancelled.")

    input("  Press ENTER to continue...")


def show_help():
    """Display help information."""
    clear_screen()
    print_header()
    print("\n  📖 HELP & DOCUMENTATION")
    print_divider("=")

    help_text = """
  QUICK START
  -----------
  1. Run Phase 1 to collect expert metadata
  2. Run Phase 2 to extract expert details
  3. Repeat for Facilities (3→4) and Organizations (5→6)

  EXECUTION MODES
  ---------------
  • Single Phase: Run one specific phase
  • Category Pipeline: Run both phases for a category
  • Full Pipeline: Run all 6 phases sequentially

  BROWSER SESSION
  ---------------
  The controller maintains a single browser session.
  Login once and it persists across multiple phases.

  CHECKPOINTS
  -----------
  Progress is saved automatically. If interrupted:
  - Re-run the phase
  - Choose 'Resume from checkpoint' when prompted

  FILE LOCATIONS
  --------------
  • Output data:  ./output/
  • Checkpoints:  ./checkpoints/
  • Error logs:   ./logs/

  TIPS
  ----
  • Check Status Dashboard before running phases
  • Run during off-peak hours for best performance
  • Keep the browser window visible (don't minimize)
  • Don't interact with the browser while running
    """
    print(help_text)
    input("  Press ENTER to return to menu...")


# ==========================================
# MAIN MENU
# ==========================================
def main_menu():
    """Display and handle main menu."""
    while True:
        clear_screen()
        print_header()

        # Quick status line
        browser_status = "🟢 Active" if browser.is_active() else "⚪ Inactive"
        login_status = "🔓 Logged In" if browser.is_logged_in else "🔒 Not Logged In"

        print(f"\n  Browser: {browser_status}  |  {login_status}")
        print_divider("=")

        print("\n  📋 MAIN MENU")
        print_divider()

        print("\n  EXECUTION:")
        print("    [1] Run Single Phase")
        print("    [2] Run Category Pipeline (Experts/Facilities/Organizations)")
        print("    [3] Run Full Pipeline (All 6 Phases)")

        print("\n  STATUS:")
        print("    [4] View Status Dashboard")
        print("    [5] Browser Session Management")

        print("\n  MAINTENANCE:")
        print("    [6] Clean Checkpoints")
        print("    [7] Clean All Data")

        print("\n  OTHER:")
        print("    [8] Help & Documentation")
        print("    [0] Exit")

        print_divider()
        choice = input("  Enter choice: ").strip()

        if choice == '1':
            run_single_phase()
        elif choice == '2':
            run_category_pipeline()
        elif choice == '3':
            run_full_pipeline()
        elif choice == '4':
            show_status_dashboard()
        elif choice == '5':
            browser_management_menu()
        elif choice == '6':
            clean_checkpoints()
        elif choice == '7':
            clean_all_data()
        elif choice == '8':
            show_help()
        elif choice == '0':
            exit_program()
            break
        else:
            print("  Invalid choice. Please try again.")
            time.sleep(1)


def browser_management_menu():
    """Browser session management submenu."""
    clear_screen()
    print_header()
    print("\n  🌐 BROWSER SESSION MANAGEMENT")
    print_divider("=")

    browser_status = "🟢 Active" if browser.is_active() else "⚪ Inactive"
    login_status = "🔓 Logged In" if browser.is_logged_in else "🔒 Not Logged In"

    print(f"\n  Current Status: {browser_status}  |  {login_status}")

    if browser.current_page:
        print(f"  Current Page: {browser.current_page[:50]}...")

    print_divider()

    print("\n    [1] Start Browser")
    print("    [2] Login to OCIP Portal")
    print("    [3] Navigate to Experts Page")
    print("    [4] Navigate to Facilities Page")
    print("    [5] Navigate to Organizations Page")
    print("    [6] Close Browser")
    print("\n    [0] Back to Main Menu")

    print_divider()
    choice = input("  Enter choice: ").strip()

    if choice == '1':
        browser.start()
        input("  Press ENTER to continue...")
    elif choice == '2':
        browser.login(force=True)
        input("  Press ENTER to continue...")
    elif choice == '3':
        if not browser.is_logged_in:
            browser.login()
        browser.navigate_to(URLS["experts"], "Experts Page")
        input("  Press ENTER to continue...")
    elif choice == '4':
        if not browser.is_logged_in:
            browser.login()
        browser.navigate_to(URLS["facilities"], "Facilities Page")
        input("  Press ENTER to continue...")
    elif choice == '5':
        if not browser.is_logged_in:
            browser.login()
        browser.navigate_to(URLS["organizations"], "Organizations Page")
        input("  Press ENTER to continue...")
    elif choice == '6':
        browser.close()
        input("  Press ENTER to continue...")


def exit_program():
    """Handle program exit."""
    clear_screen()
    print_header()
    print("\n  👋 EXITING...")
    print_divider()

    if browser.is_active():
        choice = input("\n  Close browser before exiting? (y/n): ").strip().lower()
        if choice == 'y':
            browser.close()

    print("\n  Thank you for using OCIP Extraction Pipeline!")
    print("  Goodbye.\n")


# ==========================================
# ENTRY POINT
# ==========================================
def main():
    """Main entry point."""
    try:
        main_menu()
    except KeyboardInterrupt:
        print("\n\n  ⚠️  Interrupted by user.")
        exit_program()
    except Exception as e:
        print(f"\n  ❌ Unexpected error: {e}")
        import traceback
        traceback.print_exc()
        input("\n  Press ENTER to exit...")


if __name__ == "__main__":
    main()