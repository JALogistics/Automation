import subprocess
import time

# List of scripts to run
scripts = [
    'src/WMS_Task/outbound_report.py',
    'src/RNO_Task/generate_rno_report.py',
    'src/RNO_Task/Final_rno_report.py'
]

for i, script in enumerate(scripts):
    subprocess.run(['python', script])
    if i < len(scripts) - 1:
        print(f"Finished running {script}. Waiting 20 seconds before next...")
        time.sleep(25)