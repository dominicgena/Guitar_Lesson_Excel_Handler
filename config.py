import os
# Define the root directory of the project

#region settings
activate_debug_mode = True



ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
SLEEP_INTERVAL = 10
# Define shared paths
EXCEL_FILE = os.path.join(ROOT_DIR, "Guitar-lessons.xlsm")  # Path to the main Excel file
EXCEL_SHEET = "Lesson Schedule"  # Default sheet name in the workbook
TRIGGERS_DIR = os.path.join(ROOT_DIR, "triggers")
TEMP_DIR = os.path.join(ROOT_DIR, "temp")  # Temp directory for lock files
SAVE_LOCK_FILE = os.path.join(TEMP_DIR, "autosave.lock")  # Lock file for autosave

# Autosave-specific paths
AUTOSAVE_DIR = os.path.join(ROOT_DIR, "autosave")  # Autosave directory
SAVE_TRIGGER_FILE = os.path.join(TRIGGERS_DIR, "autosave_trigger.txt")  # Trigger file for autosave
AUTOSAVE_VENV_DIR = os.path.join(AUTOSAVE_DIR, ".venv")  # Virtual environment for autosave
STATE_FILE_PATH = "triggers/statusbar-state.txt"
# Background color-specific paths
BACKGROUND_COLOR_DIR = os.path.join(ROOT_DIR, "background_color")  # Directory for background color program
BACKGROUND_COLOR_VENV_DIR = os.path.join(BACKGROUND_COLOR_DIR, ".venv-bg-color")  # Virtual environment for background color

# Define column mappings for Excel sheet
EXCEL_COLUMNS = {
    "name": "A",              # Column for student names
    "day_of_week": "C",       # Column for days of the week
    "start_time": "G",        # Column for lesson start times
    "end_time": "H",          # Column for lesson end times
    "active_status": "Q",     # Column for active/inactive status (Yes/No)
}

# Ensure required directories exist
os.makedirs(TEMP_DIR, exist_ok=True)  # Create temp directory if it doesn't exist
os.makedirs(AUTOSAVE_DIR, exist_ok=True)  # Ensure autosave directory exists
os.makedirs(BACKGROUND_COLOR_DIR, exist_ok=True)  # Ensure background color directory exists

# Workbook and Sheet Lazy Initialization
workbook = None
sheet = None

def get_workbook_and_sheet(retries=3, delay=2):
    """
    Initializes and returns the workbook and sheet, with retry logic and detailed state logging.
    """
    import xlwings as xw
    import time

    for attempt in range(1, retries + 1):
        try:
            print(f"Initializing workbook and sheet (Attempt {attempt}/{retries})...")

            # Start or attach to Excel
            if not xw.apps:
                app = xw.App(visible=True)
                print("Started a new Excel instance.")
            else:
                app = xw.apps.active
                print(f"Attached to active Excel instance. Open workbooks: {len(app.books)}")

            # Check if workbook exists
            if not os.path.exists(EXCEL_FILE):
                raise FileNotFoundError(f"Workbook not found at path: {EXCEL_FILE}")

            # Open workbook or use an existing instance
            workbook = next((wb for wb in app.books if wb.fullname == EXCEL_FILE), None)
            if workbook is None:
                print("Workbook not open. Opening now...")
                workbook = app.books.open(EXCEL_FILE)
            else:
                print(f"Workbook is already open: {workbook.fullname}")

            # List available sheets
            sheet_names = [sheet.name for sheet in workbook.sheets]
            print(f"Available sheets: {sheet_names}")

            # Access the target sheet
            if EXCEL_SHEET not in sheet_names:
                raise ValueError(f"Sheet '{EXCEL_SHEET}' not found in workbook.")
            sheet = workbook.sheets[EXCEL_SHEET]

            print(f"Successfully accessed sheet: {sheet.name}")
            return workbook, sheet

        except Exception as e:
            print(f"Error during initialization: {e}")
            if attempt < retries:
                print(f"Retrying in {delay} seconds...")
                time.sleep(delay)
            else:
                print("Max retries reached. Aborting.")
                raise


