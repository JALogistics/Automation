import os
import datetime
from pathlib import Path
import win32com.client as win32
from typing import List, Optional, Dict, Any
from loguru import logger
from glob import glob

def send_logistics_report_email(
    recipients: List[str],
    subject: Optional[str] = None,
    attachments: Optional[List[str]] = None,
    cc_recipients: Optional[List[str]] = None,
    bcc_recipients: Optional[List[str]] = None,
    report_type: str = "Logistics",
    custom_body: Optional[str] = None,
    sender_email: str = "Logistics_Reporting@jasolar.eu",
    subject_prefix: str = "JA Solar GmbH"
) -> bool:
    """
    Comprehensive function to send logistics reports via Outlook email

    Args:
        recipients: List of email addresses to send to
        subject: Optional custom subject (if None, will generate based on report type and date)
        attachments: Optional list of file paths to attach
        cc_recipients: Optional list of email addresses to CC
        bcc_recipients: Optional list of email addresses to BCC
        report_type: Type of report being sent (default: "Logistics")
        custom_body: Optional custom email body text
        sender_email: The email address to send from (default: Logistics_Reporting@jasolar.eu)
        subject_prefix: Prefix for the email subject (default: "JA Solar GmbH")

    Returns:
        bool: True if email was sent successfully, False otherwise
    """
    try:
        # Connect to Outlook
        outlook = win32.Dispatch('Outlook.Application')
        logger.info("Successfully connected to Outlook")
        
        # Create a new email
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        
        # Set sender if specified account exists in Outlook
        try:
            mail._oleobj_.Invoke(*(64209, 0, 8, 0, sender_email))
            logger.info(f"Set sender to {sender_email}")
        except Exception as e:
            logger.warning(f"Could not set sender to {sender_email}, using default account. Error: {e}")
        
        # Set recipients
        mail.To = "; ".join(recipients)
        
        # Set CC if provided
        if cc_recipients:
            mail.CC = "; ".join(cc_recipients)
            
        # Set BCC if provided
        if bcc_recipients:
            mail.BCC = "; ".join(bcc_recipients)
        
        # Set subject
        current_date = datetime.datetime.now().strftime("%d.%m.%Y")
        if subject:
            mail.Subject = subject
        else:
            mail.Subject = f"{subject_prefix} - {report_type} Report {current_date}"
        
        # Set email body (HTML formatted)
        if custom_body:
            mail.HTMLBody = custom_body
        else:
            mail.HTMLBody = f"""
            <html>
            <body style="font-family: Calibri, Arial, sans-serif; font-size: 11pt;">
                <p>Dear Team,</p>
                
                <p>Please find attached the {report_type} report dated {current_date}.</p>
                
                <p>The report contains the latest data from our system with key metrics and relevant information for your review.</p>
                
                <p>If you have any questions or need further information, please don't hesitate to contact us.</p>
                
                <p style="margin-top: 20px;">Best regards,<br>
                Logistics Reporting Team<br>
                JA Solar GmbH</p>
                
                <p style="color: #666666; font-size: 9pt; margin-top: 30px;">
                <em>This is an automated email sent from {sender_email}</em>
                </p>
            </body>
            </html>
            """
        
        # Add attachments if provided
        if attachments:
            for attachment in attachments:
                if os.path.exists(attachment):
                    mail.Attachments.Add(attachment)
                    logger.info(f"Added attachment: {attachment}")
                else:
                    logger.warning(f"Attachment file not found: {attachment}")
        
        # Send the email
        mail.Send()
        logger.success(f"Email sent successfully to {len(recipients)} recipients")
        return True
        
    except Exception as e:
        logger.error(f"Failed to send email: {e}")
        return False

if __name__ == "__main__":
    # Configure logging
    logger.add("email_sender.log", rotation="10 MB")
    
    # Test sending a sample email
    sample_recipients = ["Logistics_Reporting@jasolar.eu", "Muhammadanus@jasolar.eu"]
    sample_cc = ["Logistics_Reporting@jasolar.eu"]
    
    # Find the latest Excel file in the BMO_Reports directory
    bmo_reports_dir = r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Logistics Reporting\000_Master_Query_Reports\Automation_DB\Logistics_Reports"
    excel_files = glob(os.path.join(bmo_reports_dir, '*.xlsx'))
    
    if not excel_files:
        print(f"No Excel files found in {bmo_reports_dir}")
        test_file = None
    else:
        test_file = max(excel_files, key=os.path.getmtime)

    if test_file is None:
        print("No valid attachment found. Exiting test.")
        exit(1)
    
    # Send test email
    success = send_logistics_report_email(
        recipients=sample_recipients,
        attachments=[str(test_file)],
        cc_recipients=sample_cc
    )
    
    if success:
        print("Test email sent successfully")
    else:
        print("Failed to send test email") 