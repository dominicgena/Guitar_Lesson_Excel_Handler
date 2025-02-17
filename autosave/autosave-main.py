import time
import os
from pathlib import Path
import xlwings as xw
import logging
import sys
from datetime import datetime
import shutil
import signal
from threading import Thread, Lock, current_thread, Event
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# Adjust path to import master config
parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, parent_dir)  # Add the parent directory to the Python path
# Import the configuration file
from config import EXCEL_FILE, TEMP_DIR, SAVE_TRIGGER_FILE, activate_debug_mode, STATE_FILE_PATH

# Paths for logging and archiving
logs_folder = Path("logs")
past_logs_folder = logs_folder / "past-logs" / "Autosave"
logs_folder.mkdir(exist_ok=True)
past_logs_folder.mkdir(parents=True, exist_ok=True)

# Define the current log file path
current_log_file = logs_folder / "autosave.log"

# Logging setup
if activate_debug_mode:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        handlers=[logging.FileHandler(current_log_file, mode="w")],
    )

# Path to the lock file
lock_file_path = Path(TEMP_DIR) / "autosave.lock"

# Global debounce variables
debounce_time = 10  # Seconds
debounce_lock = Lock()
debounce_timer = None
stop_flag = Event()  # Signal for threads to exit cleanly


def log_info(message):
    if activate_debug_mode:
        logging.info(message)


def log_error(message):
    if activate_debug_mode:
        logging.error(message)


def archive_current_log():
    if not activate_debug_mode:
        return
    try:
        timestamp = datetime.now().strftime("%m-%d-%y -- %H-%M-%S")
        archived_log_file = past_logs_folder / f"Autosave {timestamp}.log"
        shutil.copy2(current_log_file, archived_log_file)
        log_info(f"Log file archived: {archived_log_file}")
    except Exception as e:
        log_error(f"Failed to archive current log file: {e}")


def signal_handler(sig, frame):
    log_info("Termination signal received. Cleaning up and archiving current log file...")
    stop_flag.set()  # Notify all threads to stop
    archive_current_log()
    sys.exit(0)


# Register signal handlers for graceful shutdown
signal.signal(signal.SIGINT, signal_handler)
signal.signal(signal.SIGTERM, signal_handler)


def is_statusbar_ready():
    try:
        if Path(STATE_FILE_PATH).exists():
            with Path(STATE_FILE_PATH).open("r") as file:
                content = file.read().strip()
            if content:
                log_info("Excel is not ready. Statusbar trigger file is not empty.")
                return False
        log_info("Excel is ready. Statusbar trigger file is empty.")
        return True
    except Exception as e:
        log_error(f"Error reading statusbar trigger file: {e}")
        return False


def wait_until_ready():
    while not stop_flag.is_set():
        if is_statusbar_ready():
            return True
        log_info("Waiting for Excel to be ready...")
        time.sleep(0.5)
    return False


def perform_save_operation():
    if stop_flag.is_set():
        return

    if not wait_until_ready():
        log_info("Save operation aborted because Excel is not ready.")
        return

    try:
        log_info("Attempting to attach to an existing Excel instance...")
        app = xw.apps.active
        workbook = next((wb for wb in app.books if wb.fullname == str(EXCEL_FILE)), None)

        if not workbook:
            log_error(f"Workbook {EXCEL_FILE} not found in active Excel instance.")
            return

        if lock_file_path.exists():
            log_info("Save operation already in progress. Skipping...")
            return

        lock_file_path.touch()
        log_info(f"Lock file created: {lock_file_path}")

        try:
            log_info(f"Saving workbook: {workbook.name}")
            workbook.save()
            log_info("Workbook saved successfully.")
        except Exception as e:
            log_error(f"Error saving workbook: {e}")
        finally:
            time.sleep(1)
            if lock_file_path.exists():
                lock_file_path.unlink()
                log_info(f"Lock file removed: {lock_file_path}")
    except Exception as e:
        log_error(f"Error during save operation: {e}")


def start_debounce_timer():
    global debounce_timer
    log_info("Starting or resetting the debounce timer.")

    def debounce_action():
        global debounce_timer
        time.sleep(debounce_time)
        with debounce_lock:
            if stop_flag.is_set():
                return
            if debounce_timer is current_thread() and is_statusbar_ready():
                log_info("Debounce timer elapsed. Performing save operation.")
                perform_save_operation()
                debounce_timer = None
            else:
                log_info("Debounce timer elapsed but conditions not met.")

    with debounce_lock:
        if debounce_timer:
            log_info("Stopping previous debounce timer.")
            debounce_timer = None

        debounce_timer = Thread(target=debounce_action, daemon=True)
        debounce_timer.start()


class TriggerFileHandler(FileSystemEventHandler):
    def __init__(self, trigger_path):
        super().__init__()
        self.trigger_path = Path(trigger_path)
        self.was_empty = self._is_file_empty()

    def _is_file_empty(self):
        try:
            with self.trigger_path.open("r") as file:
                content = file.read().strip()
            return len(content) == 0
        except FileNotFoundError:
            return True

    def on_modified(self, event):
        if Path(event.src_path) == self.trigger_path:
            is_now_empty = self._is_file_empty()
            if self.was_empty and not is_now_empty:
                log_info("Trigger file transitioned from empty to non-empty. Starting debounce timer.")
                start_debounce_timer()
            self.was_empty = is_now_empty


def monitor_trigger_file():
    trigger_path = Path(SAVE_TRIGGER_FILE)
    trigger_dir = trigger_path.parent

    event_handler = TriggerFileHandler(trigger_path)
    observer = Observer()
    observer.schedule(event_handler, str(trigger_dir), recursive=False)

    try:
        log_info("Starting trigger file monitor...")
        observer.start()
        while not stop_flag.is_set():
            time.sleep(0.1)
    except KeyboardInterrupt:
        log_info("Stopping trigger file monitor...")
    finally:
        observer.stop()
        observer.join()


if __name__ == "__main__":
    log_info("Autosave script started.")
    monitor_trigger_file()
