# WMS Report Download Automation

Python script to automate downloading the Overseas Shipment Report from the WMS system.

## Prerequisites

1. **Python 3.8+** installed on your system
2. **Google Chrome** browser installed
3. **ChromeDriver** - The script uses Selenium which requires ChromeDriver

## Installation

### Step 1: Install Python Dependencies

```bash
pip install -r requirements.txt
```

### Step 2: Install ChromeDriver

**Option A: Automatic (Recommended)**
```bash
pip install webdriver-manager
```

Then modify the script to use:
```python
from webdriver_manager.chrome import ChromeDriverManager
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
```

**Option B: Manual**
1. Download ChromeDriver from: https://chromedriver.chromium.org/downloads
2. Match the version with your Chrome browser version
3. Add ChromeDriver to your system PATH

## Usage

Simply run the script:

```bash
python download_wms_report.py
```

## What the Script Does

1. Logs into the WMS at `https://jasolar.56mada.com/oauth/login`
2. Uses credentials: Username `EU000009`, Password `JA@wms1234!`
3. Clicks the "online" option
4. Navigates to the Outbound Report page
5. Selects all organizations
6. Clicks the query button
7. Waits 30 seconds for data to load
8. Exports the report

## Configuration

You can modify the following variables in `download_wms_report.py`:

- `LOGIN_URL`: Login page URL
- `REPORT_URL`: Report page URL
- `USERNAME`: Login username
- `PASSWORD`: Login password
- `DOWNLOAD_DIR`: Where downloaded files are saved (default: `./downloads`)

## Output

Downloaded reports will be saved in the `downloads` folder in the same directory as the script.

## Troubleshooting

**Issue: ChromeDriver version mismatch**
- Solution: Install webdriver-manager (see Step 2, Option A)

**Issue: Elements not found**
- Solution: The website structure may have changed. You may need to update the selectors in the script.

**Issue: Timeout errors**
- Solution: Increase wait times in the script if you have a slow internet connection.

## Running Headless (No Browser Window)

To run without opening a visible browser window, uncomment this line in the script:

```python
chrome_options.add_argument("--headless")
```

## Scheduling Automatic Downloads

**Windows (Task Scheduler):**
1. Open Task Scheduler
2. Create a new task
3. Set trigger (e.g., daily at 9 AM)
4. Set action: Run `python C:\path\to\download_wms_report.py`

**Linux/Mac (Cron):**
```bash
crontab -e
# Add: 0 9 * * * /usr/bin/python3 /path/to/download_wms_report.py
```
