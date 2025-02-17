import time
import sys
import os
from pathlib import Path
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# Add the parent directory to the Python path for master config import
parent_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "../.."))
sys.path.insert(0, parent_dir)
# Import master config
from config import SAVE_TRIGGER_FILE, activate_debug_mode


class TriggerFileHandler(FileSystemEventHandler):
    """
    Custom handler for detecting changes from an empty to a non-empty state.
    """
    def __init__(self, trigger_path):
        super().__init__()
        self.trigger_path = Path(trigger_path)
        self.was_empty = self._is_file_empty()  # Track initial state

    def _is_file_empty(self):
        """
        Check if the trigger file is currently empty.
        """
        try:
            with self.trigger_path.open("r") as file:
                content = file.read().strip()
            return len(content) == 0
        except FileNotFoundError:
            return True  # Treat missing file as empty

    def on_modified(self, event):
        """
        Called when a file is modified. Triggers only when transitioning from empty to non-empty.
        """
        if Path(event.src_path) == self.trigger_path:
            is_now_empty = self._is_file_empty()

            # Trigger only if transitioning from empty to non-empty
            if self.was_empty and not is_now_empty:
                print("Change detected: File transitioned from empty to non-empty!")
            self.was_empty = is_now_empty  # Update the state


def monitor_trigger_file():
    """
    Monitors the trigger file for changes using watchdog.
    """
    trigger_path = Path(SAVE_TRIGGER_FILE)
    trigger_dir = trigger_path.parent

    # Debug: Output paths being monitored
    print(f"Monitoring file: {trigger_path}")
    print(f"In directory: {trigger_dir}")

    event_handler = TriggerFileHandler(trigger_path)
    observer = Observer()
    observer.schedule(event_handler, str(trigger_dir), recursive=False)

    try:
        print("Starting file monitor...")
        observer.start()
        while True:
            time.sleep(1)  # Keep the script running
    except KeyboardInterrupt:
        print("Stopping file monitor...")
        observer.stop()
    observer.join()


if __name__ == "__main__":
    if(activate_debug_mode):
        monitor_trigger_file()
    else:
        print(f"To view debug outputs, please set activate_debug_mode in config.py to True")

