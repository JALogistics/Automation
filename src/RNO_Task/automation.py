import subprocess
import time

# List of scripts to run
scripts = [
    'src/RNO_Task/Not_Released.py',
    'src/RNO_Task/stock_report.py'
]

for i, script in enumerate(scripts):
    subprocess.run(['python', script])
    if i < len(scripts) - 1:
        print(f"Finished running {script}. Waiting 20 seconds before next...")
        time.sleep(20)