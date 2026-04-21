import time
import csv
import os
from datetime import datetime
from pathlib import Path

class TimeLogger:
    """Central time logging utility that records execution times to CSV."""
    
    def __init__(self, log_file="output_data/timing_log.csv"):
        self.log_file = log_file
        self.start_time = None
        self.tasks = []
        
    def start(self):
        """Start overall timing."""
        self.start_time = time.time()
        return self.start_time
    
    def log_task(self, task_name, duration=None):
        """Log a task with its duration."""
        if duration is None and self.start_time:
            duration = time.time() - self.start_time
        self.tasks.append({
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "task": task_name,
            "duration_sec": round(duration, 2)
        })
    
    def save(self):
        """Save all logged tasks to CSV."""
        # Ensure directory exists
        os.makedirs(os.path.dirname(self.log_file), exist_ok=True)
        
        # Check if file exists to determine if we need header
        file_exists = os.path.isfile(self.log_file)
        
        with open(self.log_file, 'a', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=["timestamp", "task", "duration_sec"])
            if not file_exists:
                writer.writeheader()
            writer.writerows(self.tasks)
        
        return self.log_file
    
    def __enter__(self):
        self.start()
        return self
    
    def __exit__(self, *args):
        self.log_task("Total Execution", time.time() - self.start_time)
        self.save()