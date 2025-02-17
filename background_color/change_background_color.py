import os
import time
import sys
import tkinter as tk
from tkinter import messagebox
import openpyxl
from config import SAVE_LOCK_FILE, EXCEL_FILE, EXCEL_COLUMNS


# Function to display an error dialog box
def show_error_dialog(message):
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    messagebox.showerror("Error", message)
    root.destroy()


# Function to check if autosave is in progress
def is_autosave_in_progress(lock_file):
    return os.path.exists(lock_file)


# Function to load the latest lesson information
def load_lesson_data():
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        sheet = wb.active

        lesson_data = []
        for row in sheet.iter_rows(min_row=2, values_only=True):  # Assuming first row is headers
            data = {
                "name": row[ord(EXCEL_COLUMNS["name"]) - ord('A')],
                "day_of_week": row[ord(EXCEL_COLUMNS["day_of_week"]) - ord('A')],
                "start_time": row[ord(EXCEL_COLUMNS["start_time"]) - ord('A')],
                "end_time": row[ord(EXCEL_COLUMNS["end_time"]) - ord('A')],
                "active_status": row[ord(EXCEL_COLUMNS["active_status"]) - ord('A')],
            }
            lesson_data.append(data)
        return lesson_data
    except Exception as e:
        show_error_dialog(f"Error loading Excel data: {str(e)}")
        sys.exit(1)


# Function to apply background color changes (mock implementation)
def apply_background_colors(lesson_data):
    print("Applying background colors to lessons...")
    # Placeholder for actual implementation using libraries like `openpyxl` or `xlwings`.


# Main logic to manage background color updates
def update_background_colors():
    retry_count = 0
    max_retries = 10
    retry_delay = 2  # Seconds

    while True:
        if is_autosave_in_progress(SAVE_LOCK_FILE):
            print(f"Autosave in progress. Retrying in {retry_delay} seconds...")
            time.sleep(retry_delay)
            retry_count += 1

            if retry_count >= max_retries:
                show_error_dialog(
                    f"The autosave lock file has existed for too long.\n"
                    f"Please check the autosave process and try again."
                )
                sys.exit(1)
        else:
            break

    # Load the latest lesson data and apply colors
    lesson_data = load_lesson_data()
    apply_background_colors(lesson_data)
    print("Background colors updated successfully.")


if __name__ == "__main__":
    update_background_colors()
