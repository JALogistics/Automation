import os
import shutil
from pathlib import Path
import logging

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

class CMRFileExtractor:
    def __init__(self):
        # List to store search keys
        self.search_keys = []
        
        # Define the destination folder for CMR files
        self.cmr_destination = Path(r"C:\Users\DeepakSureshNidagund\Downloads\Reporting Application\Automation\automation\tests\CMRs")
        
        # Create the destination folder if it doesn't exist
        self.cmr_destination.mkdir(parents=True, exist_ok=True)

    def add_search_key(self, key):
        """Add a search key to the list."""
        if key and isinstance(key, str):
            self.search_keys.append(key.lower())
            logger.info(f"Added search key: {key}")
        else:
            logger.warning("Invalid search key provided")

    def find_pdf_files(self, start_path):
        """Find PDF files that match any of the search keys."""
        try:
            for root, _, files in os.walk(start_path):
                for file in files:
                    if file.lower().endswith('.pdf'):
                        file_path = Path(root) / file
                        file_name_lower = file.lower()
                        
                        # Check if any search key matches the filename
                        for key in self.search_keys:
                            if key in file_name_lower:
                                self.copy_file_to_destination(file_path)
                                break
        except Exception as e:
            logger.error(f"Error while searching for PDF files: {str(e)}")

    def copy_file_to_destination(self, source_file):
        """Copy matched PDF file to the destination folder."""
        try:
            destination_file = self.cmr_destination / source_file.name
            
            # Check if file already exists in destination
            if destination_file.exists():
                logger.warning(f"File already exists in destination: {source_file.name}")
                return
            
            shutil.copy2(source_file, destination_file)
            logger.info(f"Successfully copied: {source_file.name}")
            
        except Exception as e:
            logger.error(f"Error copying file {source_file}: {str(e)}")

def main():
    # Create an instance of CMRFileExtractor
    extractor = CMRFileExtractor()
    
    # Add your search keys here
    search_keys = [
        "2409270021",
        "2411070042",
        "2411070045",
        "2411070049",
        "2411070048",
        "2411070050",
        "2411110015",
        "2411110016",
        "2411110017"


        # Add your search keys here, for example:
        # "CMR-2024",
        # "Request-",
        # Add more as needed
    ]
    
    # Add the search keys to the extractor
    for key in search_keys:
        extractor.add_search_key(key)
    
    # Define the root directory to start searching from
    root_directory = r"C:\Users\DeepakSureshNidagund\OneDrive - JA Solar GmbH\Shared Documents - Supplier Management"# Modify this path as needed
    
    # Start the search process
    logger.info("Starting PDF file search...")
    extractor.find_pdf_files(root_directory)
    logger.info("Search completed!")

if __name__ == "__main__":
    main()
