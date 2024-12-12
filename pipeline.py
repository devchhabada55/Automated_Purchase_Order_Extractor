import os
import logging
import argparse
from typing import Optional, List

# Import the individual module functions
from gmailreader import extract_emails_to_excel
from email_classification import classify_emails_in_file
from data_extraction import POExtractor

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s: %(message)s',
    handlers=[
        logging.FileHandler('po_extraction_pipeline.log'),
        logging.StreamHandler()
    ]
)

def validate_date(date_str: str) -> bool:
    """
    Validate date string format.
    
    Args:
        date_str (str): Date string to validate
    
    Returns:
        bool: True if date format is valid, False otherwise
    """
    try:
        from datetime import datetime
        datetime.strptime(date_str, '%Y-%m-%d')
        return True
    except ValueError:
        return False

def check_file_type(file_path: str) -> str:
    """
    Determine the file type based on file extension.
    
    Args:
        file_path (str): Path to the file
    
    Returns:
        str: File type (excel, pdf, image, word, etc.)
    """
    if not os.path.exists(file_path):
        logging.error(f"File not found: {file_path}")
        return 'unknown'
    
    # Get file extension
    _, ext = os.path.splitext(file_path)
    ext = ext.lower()
    
    file_type_map = {
        '.xlsx': 'excel',
        '.xls': 'excel',
        '.pdf': 'pdf',
        '.jpg': 'image',
        '.jpeg': 'image',
        '.png': 'image',
        '.tiff': 'image',
        '.bmp': 'image',
        '.docx': 'word',
        '.doc': 'word'
    }
    
    return file_type_map.get(ext, 'unknown')

def run_pipeline(
    start_date: str, 
    end_date: str, 
    input_file: Optional[str] = None,
    classification_model: str = 'openai'
):
    """
    Run the complete Purchase Order extraction pipeline.
    
    Args:
        start_date (str): Start date for email extraction in 'YYYY-MM-DD' format
        end_date (str): End date for email extraction in 'YYYY-MM-DD' format
        input_file (str, optional): Specific input file to process
        classification_model (str, optional): Model to use for classification
    """
    try:
        # Validate date inputs
        if not (validate_date(start_date) and validate_date(end_date)):
            raise ValueError("Invalid date format. Use YYYY-MM-DD.")

        # Create required directories
        os.makedirs("attachments", exist_ok=True)
        
        # Step 1: Email Extraction
        logging.info(f"Starting email extraction from {start_date} to {end_date}")
        extract_emails_to_excel(start_date, end_date)
        extracted_file = 'emails_data-testcase.xlsx'
        logging.info("Email extraction completed successfully")

        # Step 2: Email Classification
        # Allow optional input file and classification model specification
        input_file = input_file or extracted_file
        logging.info(f"Starting email classification using {classification_model}")
        
        # Validate input file
        input_file_type = check_file_type(input_file)
        if input_file_type != 'excel':
            raise ValueError(f"Invalid input file type: {input_file_type}. Expected Excel file.")
        
        # Prepare output classification file
        output_classified_file = 'classified_emails-test-case.xlsx'
        
        # Call classification with more flexibility
        if classification_model.lower() == 'openai':
            from email_classification import classify_emails_in_file
            classify_emails_in_file(
                input_file, 
                output_classified_file, 
                os.environ.get("OPENAI_API_KEY", "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")
            )
        else:
            raise ValueError(f"Unsupported classification model: {classification_model}")
        
        logging.info("Email classification completed successfully")

        # Step 3: PO Details Extraction
        logging.info("Starting PO details extraction")
        po_extractor = POExtractor(
            input_excel=output_classified_file, 
            attachments_folder='attachments', 
            output_json='po_extracted-testcase.json'
        )
        po_extractor.process_emails()
        logging.info("PO details extraction completed successfully")

        # Verification of output files
        output_files = [
            extracted_file,
            output_classified_file,
            'po_extracted-testcase.json'
        ]

        logging.info("Verifying output files:")
        for file in output_files:
            if os.path.exists(file):
                logging.info(f"{file} - Found (Size: {os.path.getsize(file)} bytes)")
            else:
                logging.warning(f"{file} - Not found!")

        print("Purchase Order Extraction Pipeline completed successfully!")

    except Exception as e:
        logging.error(f"Pipeline execution failed: {str(e)}")
        raise

def main():
    """
    Main function to run the pipeline with command-line argument parsing.
    """
    parser = argparse.ArgumentParser(description="Purchase Order Extraction Pipeline")
    parser.add_argument('--start-date', required=True, help='Start date (YYYY-MM-DD)')
    parser.add_argument('--end-date', required=True, help='End date (YYYY-MM-DD)')
    parser.add_argument('--input-file', help='Optional specific input file to process')
    parser.add_argument('--model', choices=['openai', 'local'], default='openai', 
                        help='Classification model to use')

    args = parser.parse_args()

    # Run pipeline with parsed arguments
    run_pipeline(
        args.start_date, 
        args.end_date, 
        args.input_file, 
        args.model
    )

if __name__ == "__main__":
    main()