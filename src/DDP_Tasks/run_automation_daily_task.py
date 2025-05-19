import subprocess
import time

# List of scripts to run
scripts = [
    'src/tasks/daily-data-transfer.py',
    'src/tasks/create-consolidated-report.py',
    'src/tasks/logistics-report.py',
    'src/tasks/bmo-report.py',
    'src/tasks/mail-for-bmo-reporting.py',
    'src/tasks/mail-for-logisitcs-reporting.py'
]

for i, script in enumerate(scripts):
    subprocess.run(['python', script])
    if i < len(scripts) - 1:
        print(f"Finished running {script}. Waiting 20 seconds before next...")
        time.sleep(20)