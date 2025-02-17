import time
from datetime import datetime
from threading import Thread
from watchdog.observers import Observer
from NextLesson import fetch_data, excel_time_to_datetime, FindNextLesson, get_next_lesson
from config import SAVE_LOCK_FILE
import os


class CurrentLesson(FindNextLesson):
    """
    Handles determining the lesson currently in session by inheriting methods from NextLesson.
    Synchronizes with NextLesson to ensure data consistency.
    """
    def __init__(self):
        super().__init__()
        self.current_lesson = None
        self.lessons = fetch_data()  # Initialize the shared array of lessons
        self.lock_file_thread = Thread(target=self._watch_lock_file, daemon=True)
        self.lock_file_thread.start()

    def _watch_lock_file(self):
        """
        Watches the SAVE_LOCK_FILE and reinitializes lesson data when the file is detected.
        """
        observer = Observer()
        observer.schedule(self, os.path.dirname(SAVE_LOCK_FILE), recursive=False)

        try:
            print("Starting lock file monitor...")
            observer.start()
            observer.join()  # Keep the thread running indefinitely
        except KeyboardInterrupt:
            print("Stopping lock file monitor...")
            observer.stop()
        observer.join()

    def get_current_lesson(self):
        """
        Determines the lesson currently in session, if any.
        If a lesson is in session, returns its details; otherwise, returns None.
        """
        now = datetime.now()

        for lesson in self.lessons:
            start_time = excel_time_to_datetime(lesson["start_time"], lesson["weekday"])
            end_time = excel_time_to_datetime(lesson["end_time"], lesson["weekday"])

            # Check if the current time is within the lesson's start and end times
            if start_time <= now < end_time:
                return lesson

        return None

    def monitor_current_lesson(self):
        """
        Main loop to monitor and handle the current lesson.
        Synchronizes with NextLesson's array and handles lock file detection.
        """
        while True:
            # Find the current lesson
            self.current_lesson = self.get_current_lesson()

            if self.current_lesson:
                # Current lesson is in session
                end_time = excel_time_to_datetime(self.current_lesson["end_time"], self.current_lesson["weekday"])
                print(f"Current lesson in session: {self.current_lesson['name']} on {self.current_lesson['weekday']}. Ends at {end_time.strftime('%I:%M %p')}.")

                # Sleep until the lesson ends or a lock file is detected
                while datetime.now() < end_time:
                    if self.lock_detected:
                        print("Lock file detected during lesson session. Reinitializing data...")
                        self.lessons = fetch_data()
                        self.lock_detected = False
                        break
                    time.sleep(1)  # Sleep in 1-second increments

            else:
                # No lesson in session
                next_lesson = get_next_lesson()

                if next_lesson:
                    next_lesson_start = excel_time_to_datetime(next_lesson["start_time"], next_lesson["weekday"])

                    # Calculate sleep duration until the next lesson starts
                    sleep_duration = max(0, (next_lesson_start - datetime.now()).total_seconds())
                    print(f"No lesson in session. Sleeping for {sleep_duration:.2f} seconds until the next lesson starts.")

                    # Sleep until the next lesson starts, breaking if a lock file is detected
                    start_time = datetime.now()
                    while (datetime.now() - start_time).total_seconds() < sleep_duration:
                        if self.lock_detected:
                            print("Lock file detected during sleep period. Reinitializing data...")
                            self.lessons = fetch_data()
                            self.lock_detected = False
                            break
                        time.sleep(1)

                else:
                    print("No upcoming lessons. Rechecking in 10 seconds...")
                    time.sleep(10)  # Check for lessons again after 10 seconds


if __name__ == "__main__":
    current_lesson_tracker = CurrentLesson()
    current_lesson_tracker.monitor_current_lesson()
