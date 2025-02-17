import time
import keyboard
import win32com.client
import win32gui
import logging
import os
from pathlib import Path
from datetime import datetime
import shutil

# Paths for logging and archiving
logs_folder = Path("logs")
past_logs_folder = logs_folder / "past-logs" / "fix-esc-exc"
current_log_file = logs_folder / "escape_key_behavior.log"

# Ensure directories exist
logs_folder.mkdir(exist_ok=True)
past_logs_folder.mkdir(parents=True, exist_ok=True)

# Logging setup
logging.basicConfig(
    filename=current_log_file,
    filemode="w",  # Overwrite log each session
    format="%(asctime)s - %(levelname)s - %(message)s",
    level=logging.INFO
)

logging.info("Listening for the Escape key...")

def archive_current_log():
    """Archive the current log file to the past-logs directory."""
    try:
        timestamp = datetime.now().strftime("%m-%d-%y -- %H-%M-%S")
        archived_log_file = past_logs_folder / f"Escape {timestamp}.log"
        shutil.copy2(current_log_file, archived_log_file)
        logging.info(f"Log file archived: {archived_log_file}")
    except Exception as e:
        logging.error(f"Failed to archive log file: {e}")

def is_excel_active():
    """Check if the active window belongs to Excel."""
    try:
        active_window = win32gui.GetWindowText(win32gui.GetForegroundWindow())
        if "Excel" in active_window:
            return True
    except Exception as e:
        logging.error(f"Error checking active window: {e}")
    return False

def handle_escape_key():
    try:
        # Connect to Excel
        excel = win32com.client.GetObject(None, "Excel.Application")
        workbook = excel.ActiveWorkbook
        sheet = excel.ActiveSheet

        active_cell = sheet.Application.ActiveCell
        current_row = active_cell.Row
        current_col = active_cell.Column

        if current_row == 1:
            # Move down one, then back up
            logging.info("Row 1 detected. Moving down and back up.")
            sheet.Cells(current_row + 1, current_col).Select()
            time.sleep(0.1)  # Slight delay for movement
            sheet.Cells(current_row, current_col).Select()
        else:
            # Move up one, then back down
            logging.info("Not in row 1. Moving up and back down.")
            sheet.Cells(current_row - 1, current_col).Select()
            time.sleep(0.1)  # Slight delay for movement
            sheet.Cells(current_row, current_col).Select()

    except Exception as e:
        logging.error(f"Error handling escape key: {e}")

def main():
    logging.info("Program started.")
    escape_pressed = False  # Track if Escape key was already pressed
    try:
        while True:
            # Check if Excel is the active window
            if not is_excel_active():
                time.sleep(0.1)
                continue

            # Listen for the Escape key
            if keyboard.is_pressed("esc"):
                if not escape_pressed:  # Only handle the first press
                    logging.info("Escape key pressed.")
                    handle_escape_key()
                    escape_pressed = True
            else:
                escape_pressed = False  # Reset when Escape is released

            time.sleep(0.1)  # Avoid high CPU usage
    except KeyboardInterrupt:
        logging.info("Program interrupted by user. Exiting.")
    finally:
        archive_current_log()

if __name__ == "__main__":
    main()
