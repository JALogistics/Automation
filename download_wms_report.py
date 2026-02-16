"""
WMS Report Download Automation Script
Logs into the WMS system and downloads the Overseas Shipment Report
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
import time
import os
from datetime import datetime

# Configuration
LOGIN_URL = "https://jasolar.56mada.com/oauth/login"
REPORT_URL = "https://jasolar.56mada.com/winv-t3/report/outbound-report"
USERNAME = "EU000009"
PASSWORD = "JA@wms1234!"
DOWNLOAD_DIR = os.path.join(os.getcwd(), "downloads")

def setup_driver():
    """Configure and return Chrome WebDriver with download settings"""
    
    # Create downloads directory if it doesn't exist
    if not os.path.exists(DOWNLOAD_DIR):
        os.makedirs(DOWNLOAD_DIR)
    
    chrome_options = Options()
    
    # Set download preferences
    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)
    
    # Uncomment the line below to run headless (without opening browser window)
    # chrome_options.add_argument("--headless")
    
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    
    # Automatically install and use the correct ChromeDriver version
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver

def login(driver):
    """Login to WMS system"""
    print(f"[{datetime.now().strftime('%H:%M:%S')}] Navigating to login page...")
    driver.get(LOGIN_URL)
    
    wait = WebDriverWait(driver, 20)
    
    try:
        # Wait for username field and enter credentials
        print(f"[{datetime.now().strftime('%H:%M:%S')}] Entering credentials...")
        username_field = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='text'], input[name='username'], input[placeholder*='user']"))
        )
        username_field.clear()
        username_field.send_keys(USERNAME)
        
        # Enter password
        password_field = driver.find_element(By.CSS_SELECTOR, "input[type='password']")
        password_field.clear()
        password_field.send_keys(PASSWORD)
        
        # Click the "online" button/checkbox if it exists
        try:
            online_element = driver.find_element(By.XPATH, "//*[contains(text(), 'online') or contains(text(), 'Online') or contains(@value, 'online')]")
            online_element.click()
            print(f"[{datetime.now().strftime('%H:%M:%S')}] Clicked 'online' option")
        except NoSuchElementException:
            print(f"[{datetime.now().strftime('%H:%M:%S')}] No 'online' option found, proceeding...")
        
        # Click login button
        login_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Login') or contains(text(), 'login') or contains(text(), '登录') or @type='submit']")
        login_button.click()
        print(f"[{datetime.now().strftime('%H:%M:%S')}] Login button clicked")
        
        # Wait for successful login (adjust selector as needed)
        time.sleep(3)
        print(f"[{datetime.now().strftime('%H:%M:%S')}] Login successful!")
        
    except TimeoutException:
        print(f"[{datetime.now().strftime('%H:%M:%S')}] ERROR: Timeout while trying to login")
        raise
    except Exception as e:
        print(f"[{datetime.now().strftime('%H:%M:%S')}] ERROR during login: {str(e)}")
        raise

def select_all_organizations(driver):
    """Select all organizations in the dropdown"""
    wait = WebDriverWait(driver, 20)
    
    try:
        print(f"[{datetime.now().strftime('%H:%M:%S')}] Looking for Organization dropdown...")
        
        # Look for the organization dropdown - try multiple selectors
        org_selectors = [
            "//input[@placeholder='Organization' or contains(@placeholder, 'Organiza')]",
            "//div[contains(@class, 'el-select')]//input[contains(@placeholder, 'Organiza')]",
            "//label[contains(text(), 'Organiza')]/..//input",
            "//*[contains(text(), 'Organiza')]/following::input[1]"
        ]
        
        org_field = None
        for selector in org_selectors:
            try:
                org_field = driver.find_element(By.XPATH, selector)
                break
            except NoSuchElementException:
                continue
        
        if org_field:
            org_field.click()
            time.sleep(1)
            print(f"[{datetime.now().strftime('%H:%M:%S')}] Organization dropdown opened")
            
            # Try to select all - look for "Select All" option or check all checkboxes
            try:
                # Try to find "Select All" option
                select_all = driver.find_element(By.XPATH, "//*[contains(text(), 'Select all') or contains(text(), 'Select All') or contains(text(), '全选')]")
                select_all.click()
                print(f"[{datetime.now().strftime('%H:%M:%S')}] Selected 'Select All' option")
            except NoSuchElementException:
                # If no "Select All", try to check all checkboxes
                checkboxes = driver.find_elements(By.CSS_SELECTOR, ".el-checkbox__input:not(.is-checked), .el-checkbox")
                for checkbox in checkboxes[:10]:  # Select first 10 organizations if no "Select All"
                    try:
                        checkbox.click()
                        time.sleep(0.2)
                    except:
                        pass
                print(f"[{datetime.now().strftime('%H:%M:%S')}] Selected multiple organizations")
            
            # Close dropdown
            time.sleep(1)
            driver.find_element(By.TAG_NAME, "body").click()
        else:
            print(f"[{datetime.now().strftime('%H:%M:%S')}] WARNING: Could not find Organization dropdown")
            
    except Exception as e:
        print(f"[{datetime.now().strftime('%H:%M:%S')}] ERROR while selecting organizations: {str(e)}")
        # Continue anyway as this might not be critical

def query_report(driver):
    """Click the query button to load report data"""
    wait = WebDriverWait(driver, 20)
    
    try:
        print(f"[{datetime.now().strftime('%H:%M:%S')}] Looking for Query button...")
        
        # Look for query button - try multiple selectors
        query_selectors = [
            "//button[contains(text(), 'query') or contains(text(), 'Query')]",
            "//button[contains(@class, 'query')]",
            "//span[text()='query']/parent::button",
            "//button//span[text()='query']"
        ]
        
        query_button = None
        for selector in query_selectors:
            try:
                query_button = driver.find_element(By.XPATH, selector)
                break
            except NoSuchElementException:
                continue
        
        if query_button:
            query_button.click()
            print(f"[{datetime.now().strftime('%H:%M:%S')}] Query button clicked")
            print(f"[{datetime.now().strftime('%H:%M:%S')}] Waiting 30 seconds for data to load...")
            time.sleep(30)
            print(f"[{datetime.now().strftime('%H:%M:%S')}] Data should be loaded now")
        else:
            print(f"[{datetime.now().strftime('%H:%M:%S')}] WARNING: Could not find Query button")
            
    except Exception as e:
        print(f"[{datetime.now().strftime('%H:%M:%S')}] ERROR while querying: {str(e)}")
        raise

def export_report(driver):
    """Click the export button to download the report"""
    wait = WebDriverWait(driver, 20)
    
    try:
        print(f"[{datetime.now().strftime('%H:%M:%S')}] Looking for Export button...")
        
        # Look for export button - try multiple selectors
        export_selectors = [
            "//button[contains(text(), 'export') or contains(text(), 'Export')]",
            "//button[contains(@class, 'export')]",
            "//span[text()='export']/parent::button",
            "//*[@class='el-icon-download']",
            "//button//span[contains(text(), 'export')]"
        ]
        
        export_button = None
        for selector in export_selectors:
            try:
                export_button = driver.find_element(By.XPATH, selector)
                break
            except NoSuchElementException:
                continue
        
        if export_button:
            export_button.click()
            print(f"[{datetime.now().strftime('%H:%M:%S')}] Export button clicked")
            print(f"[{datetime.now().strftime('%H:%M:%S')}] Waiting for download to complete...")
            time.sleep(5)
            print(f"[{datetime.now().strftime('%H:%M:%S')}] Download should be complete")
        else:
            print(f"[{datetime.now().strftime('%H:%M:%S')}] WARNING: Could not find Export button")
            
    except Exception as e:
        print(f"[{datetime.now().strftime('%H:%M:%S')}] ERROR while exporting: {str(e)}")
        raise

def main():
    """Main execution function"""
    driver = None
    
    try:
        print("="*60)
        print("WMS Report Download Automation")
        print("="*60)
        
        # Setup driver
        driver = setup_driver()
        print(f"[{datetime.now().strftime('%H:%M:%S')}] Chrome WebDriver initialized")
        print(f"[{datetime.now().strftime('%H:%M:%S')}] Downloads will be saved to: {DOWNLOAD_DIR}")
        
        # Login
        login(driver)
        
        # Navigate to report page
        print(f"[{datetime.now().strftime('%H:%M:%S')}] Navigating to report page...")
        driver.get(REPORT_URL)
        time.sleep(3)
        
        # Select all organizations
        select_all_organizations(driver)
        
        # Query the report
        query_report(driver)
        
        # Export the report
        export_report(driver)
        
        print("="*60)
        print(f"[{datetime.now().strftime('%H:%M:%S')}] SUCCESS! Report download completed")
        print(f"[{datetime.now().strftime('%H:%M:%S')}] Check the downloads folder: {DOWNLOAD_DIR}")
        print("="*60)
        
        # Keep browser open for a few seconds to ensure download completes
        time.sleep(5)
        
    except Exception as e:
        print("="*60)
        print(f"[{datetime.now().strftime('%H:%M:%S')}] ERROR: {str(e)}")
        print("="*60)
        
    finally:
        if driver:
            print(f"[{datetime.now().strftime('%H:%M:%S')}] Closing browser...")
            driver.quit()

if __name__ == "__main__":
    main()

