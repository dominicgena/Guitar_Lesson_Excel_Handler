# ORIGINAL PATH - "logs/current-NextLesson.py"

import os
import time
import sys
parent_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "../.."))
sys.path.insert(0, parent_dir)
from config import activate_debug_mode, STATE_FILE_PATH

def monitor_log_file(trigger_file_path):
    """Monitor the NextLesson log file and display its content in real time."""
    try:
        print(f"Monitoring {trigger_file_path}...\n")
        last_position = 0

        while True:
            try:
                with open(trigger_file_path, "r") as log_file:
                    # Check if the file has been cleared
                    current_size = os.stat(trigger_file_path).st_size
                    if current_size < last_position:
                        # File cleared: Reset screen and last position
                        os.system('cls' if os.name == 'nt' else 'clear')
                        print(f"Monitoring {trigger_file_path}...\n")
                        last_position = 0

                    # Seek to the last known position and read new lines
                    log_file.seek(last_position)
                    for line in log_file:
                        print(line, end="")  # Display new lines
                    last_position = log_file.tell()
            except FileNotFoundError:
                print(f"Log file {trigger_file_path} not found. Retrying...")
                time.sleep(1)  # Wait and retry if the file is missing

            time.sleep(0.1)  # Polling interval
    except KeyboardInterrupt:
        print("\nStopped monitoring.")

if __name__ == "__main__":
    if activate_debug_mode:
        monitor_log_file(STATE_FILE_PATH)
    else:
        print(f"To view debug outputs, please set activate_debug_mode in config.py to True")
