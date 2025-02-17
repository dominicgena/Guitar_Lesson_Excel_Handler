import time
from datetime import datetime, timedelta
from pathlib import Path
import logging
from config import get_workbook_and_sheet, SAVE_LOCK_FILE

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
    try:
        result = base_date + timedelta(days=excel_serial)
        logging.debug(f"Converted Excel serial '{excel_serial}' to datetime: {result}")
        return result
    except Exception as e:
        logging.error(f"Error converting Excel serial '{excel_serial}' to datetime: {e}")
        return None



def fetch_data():
    """
    Fetch all active lesson data from the Excel sheet.
    """
    try:
        logging.info("Fetching data...")
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

                # Log raw data
                logging.info(f"Row {row}: Name={name}, Weekday={weekday}, StartTimeRaw={start_time}")

                # Convert Excel serial time to Python datetime
                start_time = excel_serial_to_datetime(start_time)
                lessons.append({"name": name, "weekday": weekday, "start_time": start_time, "row": row})
                row += 1

            except Exception as e:
                logging.error(f"Error reading row {row}: {e}")
                break

        logging.info(f"Fetched {len(lessons)} lessons.")
        for lesson in lessons:
            logging.debug(f"Lesson fetched: {lesson}")
        return lessons

    except Exception as e:
        logging.error(f"Error fetching data: {e}")
        raise




def find_next_lesson(lessons):
    """
    Finds and returns the next lesson closest to the current time.
    """
    now = datetime.now()
    logging.info(f"Current time: {now}")
    logging.info(f"Today's weekday number (now.weekday()): {now.weekday()}")

    # Manual weekday mapping
    WEEKDAY_MAP = {
        "monday": 0,
        "tuesday": 1,
        "wednesday": 2,
        "thursday": 3,
        "friday": 4,
        "saturday": 5,
        "sunday": 6,
    }

    upcoming_lessons = []
    for lesson in lessons:
        try:
            # Log raw weekday before processing
            logging.debug(f"Raw weekday for '{lesson['name']}': '{lesson['weekday']}' (ASCII: {[ord(c) for c in lesson['weekday']]})")

            # Normalize and map the weekday
            normalized_weekday = lesson["weekday"].strip().lower()
            lesson_weekday = WEEKDAY_MAP.get(normalized_weekday, -1)  # Default to -1 for invalid days

            # Log the mapping result
            if lesson_weekday == -1:
                logging.error(f"Invalid weekday '{lesson['weekday']}' for lesson '{lesson['name']}'")
                continue

            logging.info(f"Lesson '{lesson['name']}' normalized weekday '{normalized_weekday}' interpreted as {lesson_weekday}")

            # Calculate days ahead
            days_ahead = (lesson_weekday - now.weekday()) % 7
            logging.info(f"DaysAhead for '{lesson['name']}': {days_ahead}")

            # Calculate lesson datetime
            lesson_date = now + timedelta(days=days_ahead)
            lesson_datetime = lesson_date.replace(
                hour=lesson["start_time"].hour,
                minute=lesson["start_time"].minute,
                second=0,
                microsecond=0
            )
            logging.info(f"Lesson '{lesson['name']}' datetime calculated as {lesson_datetime}")

            # Correct same-day handling
            if days_ahead == 0 and lesson_datetime < now:
                lesson_datetime += timedelta(days=7)

            if lesson_datetime > now:
                lesson["lesson_datetime"] = lesson_datetime
                upcoming_lessons.append(lesson)
        except Exception as e:
            logging.error(f"Error processing lesson {lesson}: {e}")

    if upcoming_lessons:
        sorted_lessons = sorted(upcoming_lessons, key=lambda l: l["lesson_datetime"])
        return sorted_lessons[0]

    return None


def monitor_lock_file(lessons):
    """
    Poll for the presence and removal of the lock file to trigger reinitialization.
    """
    logging.info("Monitoring for lock file creation and deletion...")
    lock_file_path = Path(SAVE_LOCK_FILE)

    try:
        while True:
            if lock_file_path.exists():
                logging.info(f"Lock file detected at {SAVE_LOCK_FILE}. Waiting for removal...")
                # Wait for the lock file to be removed
                while lock_file_path.exists():
                    time.sleep(0.1)  # Poll every 100ms
                logging.info("Lock file removed. Save operation completed. Reinitializing lessons array.")
                lessons[:] = fetch_data()  # Update lessons array in-place

                # Output the next lesson after reinitialization
                next_lesson = find_next_lesson(lessons)
                if next_lesson:
                    logging.info(f"Next student: {next_lesson['name']} at {next_lesson['lesson_datetime']}")
                else:
                    logging.info("No upcoming lessons found.")
            time.sleep(0.1)  # Poll every 100ms for new lock file creation
    except KeyboardInterrupt:
        logging.info("Program interrupted. Exiting...")


def main():
    logging.info("Starting NextLesson program...")

    # Initialize lessons array
    lessons = fetch_data()

    # Output the next lesson upon startup
    next_lesson = find_next_lesson(lessons)
    if next_lesson:
        logging.info(f"Next student: {next_lesson['name']} at {next_lesson['lesson_datetime']}")
    else:
        logging.info("No upcoming lessons found.")

    # Start monitoring for lock file changes
    monitor_lock_file(lessons)


if __name__ == "__main__":
    main()
