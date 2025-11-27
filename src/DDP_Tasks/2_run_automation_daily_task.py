import subprocess
import time

# List of scripts to run
scripts = [
    'src/DDP_Task/tranfer_latest_EU and RNO _stock.py',
    'src/DDP_Tasks/wh-stock.py',
    'src/DDP_Tasks/transfer_file.py',
    'src/DDP_Tasks/cdr_stock.py'
]

for i, script in enumerate(scripts):
    subprocess.run(['python', script])
    if i < len(scripts) - 1:
        print(f"Finished running {script}. Waiting 20 seconds before next...")
        time.sleep(20)