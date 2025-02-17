import time
from pathlib import Path
from watchdog.events import FileSystemEventHandler
import logging
from datetime import datetime, timedelta  # Added this import
from config import get_workbook_and_sheet, SAVE_LOCK_FILE  # Ensure SAVE_LOCK_FILE points to the correct lock file path
from tabulate import tabulate

# Logging setup
logging.basicConfig(
    format="%(asctime)s - %(message)s",
    level=logging.INFO
)

# Helper functions
def excel_serial_to_datetime(excel_serial):
    """
    Converts an Excel serial time to a Python datetime object.
    """
    base_date = datetime(1899, 12, 30)  # Excel's base date (accounting for leap year bug)
    return base_date + timedelta(days=excel_serial)


def fetch_data():
    """
    Fetch all active lesson data from the Excel sheet.
    """
    try:
        logging.info(f"Fetching data...")
        workbook, sheet = get_workbook_and_sheet()
        lessons = []

        column_map = {"name": "A", "weekday": "C", "start_time": "G"}
        row = 2  # Start from row 2 (headers are in row 1)

        while True:
            try:
                name = sheet.range(f"{column_map['name']}{row}").value
                weekday = sheet.range(f"{column_map['weekday']}{row}").value
                start_time = sheet.range(f"{column_map['start_time']}{row}").value

                if not name or str(name).strip() == "":
                    break

                # Convert Excel serial time to Python datetime
                start_time = excel_serial_to_datetime(start_time)
                lessons.append({"name": name, "weekday": weekday, "start_time": start_time, "row": row})
                row += 1

            except Exception as e:
                logging.error(f"Error reading row {row}: {e}")
                break

        logging.info(f"Fetched {len(lessons)} lessons.")
        return lessons

    except Exception as e:
        logging.error(f"Error fetching data: {e}")
        raise


# Event handler for lock file changes
class LockFileHandler(FileSystemEventHandler):
    def __init__(self):
        super().__init__()
        self.lock_detected = False
        self.lock_removed = False

    def on_created(self, event):
        if event.src_path == SAVE_LOCK_FILE:
            logging.info("Lock file detected. Save operation started.")
            self.lock_detected = True

    def on_deleted(self, event):
        if event.src_path == SAVE_LOCK_FILE:
            logging.info("Lock file removed. Save operation completed.")
            self.lock_removed = True


# Function to monitor the lock file with polling
def monitor_lock_file():
    """
    Poll for the presence and removal of the lock file to trigger reinitialization.
    """
    logging.info("Monitoring for lock file creation and deletion...")
    lock_file_path = Path(SAVE_LOCK_FILE)

    try:
        while True:
            # Check if the lock file exists
            if lock_file_path.exists():
                logging.info("Lock file detected. Save operation started.")
                # Wait for the lock file to be removed
                while lock_file_path.exists():
                    time.sleep(0.1)  # Poll every 100ms
                logging.info("Lock file removed. Save operation completed.")

                # Trigger reinitialization
                lessons = fetch_data()
                # Display lessons in a pretty table
                if lessons:
                    table = [[lesson["row"], lesson["name"], lesson["weekday"], lesson["start_time"]] for lesson in lessons]
                    headers = ["Row", "Name", "Weekday", "Start Time"]
                    print("\n" + tabulate(table, headers=headers, tablefmt="grid"))
                else:
                    logging.info("No lessons found.")
            time.sleep(0.1)  # Poll every 100ms for new lock file creation
    except KeyboardInterrupt:
        logging.info("Program interrupted. Exiting...")


if __name__ == "__main__":
    monitor_lock_file()
