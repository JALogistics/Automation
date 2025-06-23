import os
from pathlib import Path
import win32com.client as win32
from datetime import datetime
from loguru import logger
import pandas as pd

def get_latest_file(directory):
    # Get list of files in directory
    files = Path(directory).glob('*')
    
    # Get the most recent file
    latest_file = max(files, key=lambda x: x.stat().st_mtime)
    return latest_file

def send_email_with_attachment():
    # Configure logging
    logger.add("email_sender.log", rotation="10 MB")
    
    # Directory containing the files
    directory = r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Logistics Reporting\000_Master_Query_Reports\Automation_DB\Sales_RNO_Report"
    
    try:
        # Get the latest file
        latest_file = get_latest_file(directory)
        logger.info(f"Found latest file: {latest_file}")

        # Read 'Status Summary' sheet as HTML table
        try:
            status_summary_df = pd.read_excel(latest_file, sheet_name='Status Summary')
            status_summary_html = status_summary_df.to_html(index=False, border=1, classes='status-summary-table', justify='center')
        except Exception as e:
            status_summary_html = f'<p style="color:red;">Could not read Status Summary: {str(e)}</p>'
            logger.error(f"Failed to read Status Summary: {str(e)}")
        
        # Read 'Current Location Goods' sheet as HTML table
        try:
            location_goods_df = pd.read_excel(latest_file, sheet_name='Current Location Goods')
            location_goods_html = location_goods_df.to_html(index=False, border=1, classes='location-goods-table', justify='center')
        except Exception as e:
            location_goods_html = f'<p style="color:red;">Could not read Current Location Goods: {str(e)}</p>'
            logger.error(f"Failed to read Current Location Goods: {str(e)}")
        
        # Create Outlook application object
        outlook = win32.Dispatch('Outlook.Application')
        logger.info("Connected to Outlook")
        
        # Create a new mail item
        mail = outlook.CreateItem(0)  # 0 represents olMailItem
        
        # Configure email properties
        sender_email: str = "Logistics_Reporting@jasolar.eu"
        current_date = datetime.now().strftime("%d.%m.%Y")
        mail.Subject = f"Sales RNO Report - {datetime.now().strftime('%d %B %Y')}"
        
        # Email body with embedded tables
        mail.HTMLBody = f"""
            <html>
            <body style=\"font-family: Calibri, Arial, sans-serif; font-size: 11pt;\">
                <p>Dear Team,</p>
                <p><b>Status Summary Table:</b></p>
                {status_summary_html}
                <p><b>Current Location Goods Table:</b></p>
                {location_goods_html}
                <p>Please find attached the Sales RNO report dated {current_date}.</p>
                <p>The report contains the latest data from our system with key metrics and relevant information for your review.</p>
                <p>If you have any questions or need further information, please don't hesitate to contact us.</p>
                <p style=\"margin-top: 20px;\">Best regards,<br>
                Logistics Reporting Team<br>
                JA Solar GmbH</p>
                <p style=\"color: #666666; font-size: 9pt; margin-top: 30px;\">
                <em>This is an automated email sent from {sender_email}</em>
                </p>
            </body>
            </html>
            """
        
        # Add recipients
        mail.To = ";".join(["Logistics_Reporting@jasolar.eu", "Muhammadanus@jasolar.eu"])
        mail.CC = "Logistics_Reporting@jasolar.eu"
        logger.info("Email recipients configured")
        
        # Attach the file
        mail.Attachments.Add(str(latest_file))
        logger.info(f"Attached file: {latest_file.name}")
        
        # Send the email
        mail.Send()
        logger.success("Email sent successfully")
        
        print(f"Email sent successfully with attachment: {latest_file.name}")
        
    except Exception as e:
        error_msg = f"An error occurred: {str(e)}"
        logger.error(error_msg)
        print(error_msg)

if __name__ == "__main__":
    send_email_with_attachment()
