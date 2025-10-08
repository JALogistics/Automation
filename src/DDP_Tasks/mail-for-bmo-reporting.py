import os
import datetime
from pathlib import Path
import win32com.client as win32
from typing import List, Optional, Dict, Any
from loguru import logger
from glob import glob
import pandas as pd

def send_bmo_report_email(
    sender_email: str = "Logistics_Reporting@jasolar.eu",
    recipients: List[str] = ["Logistics_Reporting@jasolar.eu", "Muhammadanus@jasolar.eu","leizhang@jasolar.eu"],
    cc_recipients: Optional[List[str]] = ["Logistics_Reporting@jasolar.eu"],
    bcc_recipients: Optional[List[str]] = None,
    report_directory: str = r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Logistics Reporting\000_Master_Query_Reports\Automation_DB\BMO_Reports",
    custom_body: Optional[str] = None
) -> bool:
    """
    Send BMO report email with the latest Excel file from the specified directory.
    
    Args:
        sender_email: The email address to send from
        recipients: List of email addresses to send to
        cc_recipients: Optional list of email addresses to CC
        bcc_recipients: Optional list of email addresses to BCC
        report_directory: Directory containing the BMO report Excel files
        custom_body: Optional custom email body text
    
    Returns:
        bool: True if email was sent successfully, False otherwise
    """
    try:
        # Set up logging
        logger.add("email_sender.log", rotation="10 MB")
        
        # Find the latest Excel file in the BMO_Reports directory
        excel_files = glob(os.path.join(report_directory, '*.xlsx'))
        if not excel_files:
            logger.error(f"No Excel files found in {report_directory}")
            return False
            
        latest_report = max(excel_files, key=os.path.getmtime)
        
        # Connect to Outlook
        try:
            outlook = win32.Dispatch('Outlook.Application')
            logger.info("Successfully connected to Outlook")
        except Exception as e:
            logger.error(f"Failed to connect to Outlook: {e}")
            return False
            
        # Create email
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        
        # Set sender
        try:
            mail._oleobj_.Invoke(*(64209, 0, 8, 0, sender_email))
            logger.info(f"Set sender to {sender_email}")
        except Exception as e:
            logger.warning(f"Could not set sender to {sender_email}, using default account. Error: {e}")
        
        # Set recipients
        mail.To = "; ".join(recipients)
        if cc_recipients:
            mail.CC = "; ".join(cc_recipients)
        if bcc_recipients:
            mail.BCC = "; ".join(bcc_recipients)
        
        # Set subject with current date
        current_date = datetime.datetime.now().strftime("%d.%m.%Y")
        mail.Subject = f"JA Solar GmbH - Logistics Outbound Report {current_date}"
        
        # Set email body (no summary table)
        if custom_body:
            mail.HTMLBody = custom_body
        else:
            mail.HTMLBody = f"""
            <html>
            <body style=\"font-family: Calibri, Arial, sans-serif; font-size: 11pt;\">
                <p>Dear Team,</p>
                <p>Please find attached the Logistics Outbound report dated {current_date}.</p>
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
        
        # Add attachment
        if os.path.exists(latest_report):
            mail.Attachments.Add(latest_report)
            logger.info(f"Added attachment: {latest_report}")
        else:
            logger.error(f"Attachment file not found: {latest_report}")
            return False
        
        # Send email
        mail.Send()
        logger.success(f"Email sent successfully to {len(recipients)} recipients")
        return True
        
    except Exception as e:
        logger.error(f"Failed to send email: {e}")
        return False

if __name__ == "__main__":
    # Example usage
    success = send_bmo_report_email()
    if success:
        print("BMO report email sent successfully")
    else:
        print("Failed to send BMO report email") 