import subprocess
import time

# List of scripts to run
scripts = [
    'src/DDP_Tasks/daily-data-transfer.py',
    'src/DDP_Tasks/create-consolidated-report.py',
    'src/DDP_Tasks/logistics-report.py',
    'src/DDP_Tasks/bmo-report.py',
    'src/DDP_Tasks/mail-for-bmo-reporting.py',
    'src/DDP_Tasks/mail-for-logisitcs-reporting.py',
    # 'src/DDP_Tasks/wh-stock.py',
    # 'src/DDP_Tasks/transfer_file.py',
    # 'src/DDP_Tasks/cdr_stock.py'
]

for i, script in enumerate(scripts):
    subprocess.run(['python', script])
    if i < len(scripts) - 1:
        print(f"Finished running {script}. Waiting 20 seconds before next...")
        time.sleep(20)